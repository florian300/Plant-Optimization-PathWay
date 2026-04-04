import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import os
import shutil
import matplotlib.ticker as ticker
from rich import print
from tqdm import tqdm
from .optimizer import PathFinderOptimizer

class PathFinderReporter:
    def __init__(self, optimizer: PathFinderOptimizer, scenario_id: str = "DEFAULT", scenario_name: str = "Default", generate_excel: bool = True, verbose: bool = False, progress_cb=None):
        self.opt = optimizer
        self.data = optimizer.data
        self.years = optimizer.years
        self.scenario_id = scenario_id
        self.scenario_name = scenario_name
        self.generate_excel = generate_excel
        self.verbose = verbose
        self.progress_cb = progress_cb
        repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
        self.results_dir = os.path.join(repo_root, 'artifacts', 'reports', self.scenario_name)
        self.charts_data = [] # List of (title, dataframe) tuples

    def _add_scenario_label(self, fig):
        """Add a colored label box in the top-right corner with the scenario name."""
        palette = {
            'BS': '#1A5276', # dark blue
            'CT': '#1E8449', # dark green
            'LCB': '#6E2F7C' # purple
        }
        color = palette.get(self.scenario_id.upper(), '#333333')
        
        # Add text box in figure coordinates (top right)
        fig.text(0.98, 0.98, f" {self.scenario_name} ", 
                 color='white', fontsize=10, weight='bold',
                 ha='right', va='top',
                 bbox=dict(facecolor=color, alpha=0.9, edgecolor='none', boxstyle='round,pad=0.3'))

    def _save_no_data_chart(self, file_name: str, title: str, message: str):
        """Creates a placeholder chart when a reporting toggle is enabled but no data is available."""
        fig, ax = plt.subplots(figsize=(11, 7), facecolor='white')
        ax.axis('off')
        ax.text(0.5, 0.62, title, ha='center', va='center', fontsize=18, weight='bold', color='#2C3E50', transform=ax.transAxes)
        ax.text(0.5, 0.45, message, ha='center', va='center', fontsize=12, color='#555555', transform=ax.transAxes)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, file_name), dpi=300, bbox_inches='tight')
        plt.close()
        
    def generate_report(self):
        # 0. Calculate total steps for progress tracking
        toggles = self.data.reporting_toggles
        # Count active charts
        steps_total = sum(1 for k, v in vars(toggles).items() if k.startswith('chart_') and v is True)
        # Add Excel export steps (Data collection + Save)
        steps_total += 2 if self.generate_excel else 0
        steps_done = 0

        def _step():
            nonlocal steps_done
            steps_done += 1
            if self.progress_cb:
                self.progress_cb(steps_done, steps_total)

        if self.verbose:
            print(f"  [yellow][Reporter][/yellow] [STATS] Generating reports ({steps_total} steps)...")
        
        # 0. Clear and Recreate Results Directory
        results_dir = self.results_dir
        if os.path.exists(results_dir):
            try:
                shutil.rmtree(results_dir)
            except Exception as e:
                print(f"  [yellow][Reporter][/yellow] [!] Could not clear results directory: {e}")
        os.makedirs(results_dir, exist_ok=True)

        # 1. Collect Data
        if self.verbose:
            print("  [yellow][Reporter][/yellow] [SEARCH] Collecting results data...")
        _step() # Mark data collection step
        # ... (rest of data collection remains same)
        investments = []
        for t in tqdm(self.years, desc="Processing Years", leave=False, disable=not self.verbose):
            for p_id, process in self.opt.entity.processes.items():
                for t_id in process.valid_technologies:
                    if t_id == 'UP':  # UP is free continuous improvement — not an investissement
                        continue
                    if self.opt.invest_vars[(t, p_id, t_id)].varValue is not None and self.opt.invest_vars[(t, p_id, t_id)].varValue > 1e-6:
                        invested_units = self.opt.invest_vars[(t, p_id, t_id)].varValue
                        tech = self.data.technologies[t_id]
                        
                        # Calculate CAPEX scaled by fractional units
                        cap_capex = 1.0
                        if tech.capex_per_unit:
                            if tech.capex_unit == 'tCO2': cap_capex = self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                            elif 'MW' in tech.capex_unit.upper(): 
                                cap_capex = (self.opt.entity.base_consumptions.get('EN_FUEL', 0.0) * process.consumption_shares.get('EN_FUEL', 0.0)) / self.opt.entity.annual_operating_hours
                        current_capex = tech.capex_by_year.get(t, tech.capex)
                        capex_cost = (current_capex * cap_capex / process.nb_units) * invested_units
                        
                        # Calculate Reductions (negative impacts) and Additions (positive impacts)
                        impact_strings = []
                        for res_id, imp in tech.impacts.items():
                            val = imp['value']
                            if res_id == 'CO2_EM':
                                base_c = self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                            else:
                                base_c = self.opt.entity.base_consumptions.get(res_id, 0.0) * process.consumption_shares.get(res_id, 0.0)
                                
                            target_change = 0.0
                            if imp['type'] == 'variation' or imp['type'] == 'up':
                                target_change = val * base_c
                            elif imp['type'] == 'new':
                                ref_res = imp.get('ref_resource')
                                if ref_res == 'CO2_EM':
                                    base_ref_amount = self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                                elif ref_res and ref_res in self.opt.entity.base_consumptions:
                                    base_ref_amount = self.opt.entity.base_consumptions.get(ref_res, 0.0) * process.consumption_shares.get(ref_res, 0.0)
                                else:
                                    base_ref_amount = 1.0
                                    
                                if imp.get('reference', '') == 'AVOIDED' and ref_res:
                                    ref_imp = tech.impacts.get(ref_res)
                                    if not ref_imp and ref_res == 'CO2_EM':
                                        co2_keys = [k for k in tech.impacts.keys() if 'CO2' in k]
                                        if co2_keys: ref_imp = tech.impacts[co2_keys[0]]
                                    if ref_imp and ref_imp['type'] in ['variation', 'up'] and ref_imp['value'] < 0:
                                        base_ref_amount = abs(ref_imp['value']) * base_ref_amount

                                # Fallback just in case for PEM_H2 original logic if no ref_resource
                                if res_id == 'EN_ELEC' and t_id == 'PEM_H2' and not ref_res:
                                    base_ref_amount = self.opt.entity.base_consumptions.get('EN_FUEL', 0.0) * process.consumption_shares.get('EN_FUEL', 0.0) * 0.5
                                    
                                target_change = val * base_ref_amount
                                
                            # Scale the targeted baseline impact by the number of active units being invested
                            target_change = target_change * (invested_units / process.nb_units)
                                
                            if abs(target_change) > 1e-4:
                                direction = "Decrease" if target_change < 0 else "Increase"
                                impact_strings.append(f"{direction} {abs(target_change):,.2f} {res_id}")
                        
                        investments.append({
                            'Year': t, 
                            'Process': p_id, 
                            'Technology': t_id, 
                            'Units_Invested': f"{invested_units:.2f}/{process.nb_units}" if invested_units % 1 != 0 else f"{int(invested_units)}/{process.nb_units}",
                            'Capex_Euros': capex_cost,
                            'Impacts': " | ".join(impact_strings)
                        })
                    
        df_invest = pd.DataFrame(investments)
        
        # Consumptions
        cons_data = []
        for t in self.years:
            row = {'Year': t}
            for res_id in self.data.resources:
                if res_id != 'CO2_EM':
                    row[res_id] = self.opt.cons_vars[(t, res_id)].varValue
            
            # --- Granular H2: Explicitly capture PRODUCED_ON_SITE ---
            if hasattr(self.opt, 'h2_supply_vars') and t in self.opt.h2_supply_vars:
                row['EN_H2_ON_SITE'] = self.opt.h2_supply_vars[t].get('PRODUCED_ON_SITE', 0.0).varValue if hasattr(self.opt.h2_supply_vars[t].get('PRODUCED_ON_SITE', 0.0), 'varValue') else 0.0
            
            cons_data.append(row)
        df_cons = pd.DataFrame(cons_data)
        
        # Emissions
        emis_data = []
        for t in self.years:
            # Emis_vars is now strictly Direct Emissions
            direct_co2 = self.opt.emis_vars[t].varValue
            
            # Index for UP factor calculation
            yr_idx = list(self.opt.years).index(t)
            
            # Calculate Indirect Emissions for this year
            indirect_co2 = 0.0
            for r_id in self.data.resources:
                if r_id in self.data.time_series.other_emissions_factors:
                    factor = self.data.time_series.other_emissions_factors[r_id].get(t, 0.0)
                    if factor > 0:
                        cons_val = self.opt.cons_vars[(t, r_id)].varValue
                        indirect_co2 += cons_val * factor
                        
            total_co2 = direct_co2 + indirect_co2
            
            if self.opt.entity.sv_act_mode == "PI":
                fq_pct = self.data.time_series.carbon_quotas_pi.get(t, 0.0)
            else:
                fq_pct = self.data.time_series.carbon_quotas_norm.get(t, 0.0)
                
            taxed_co2 = self.opt.taxed_emis_vars[t].varValue
            tax_price = self.data.time_series.carbon_prices.get(t, 0.0)
            penalty_factor = self.data.time_series.carbon_penalties.get(t, 0.0)
            # Total cost is the sum of paid and penalized quotas
            paid_q = self.opt.paid_quota_vars[t].varValue or 0.0
            penal_q = self.opt.penalty_quota_vars[t].varValue or 0.0
            tax_cost_meuros = (paid_q * tax_price + penal_q * tax_price * (1.0 + penalty_factor)) / 1_000_000.0
            
            # Calcul des émissions évitées (par rapport aux émissions de référence directes)
            avoided_total_kt = max(0.0, (self.opt.entity.base_emissions - direct_co2) / 1000.0)
            
            # Identify CCS/CCU contribution
            captured_kt = 0.0
            for p_id, process in self.opt.entity.processes.items():
                for t_id in process.valid_technologies:
                    if t_id == 'UP': continue
                    act_var = getattr(self.opt.active_vars[(t, p_id, t_id)], 'varValue', 0.0) or 0.0
                    if act_var > 1e-6:
                        if "CCS" in t_id.upper() or "CCU" in t_id.upper():
                            tech = self.data.technologies[t_id]
                            imp = tech.impacts.get('CO2_EM')
                            if not imp:
                                co2_keys = [k for k in tech.impacts.keys() if 'CO2' in k]
                                if co2_keys: imp = tech.impacts[co2_keys[0]]
                            
                            if imp:
                                initial_val = self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                                ref_res = imp.get('ref_resource')
                                if ref_res == 'CO2_EM':
                                    base_ref_amount = initial_val
                                elif ref_res and ref_res in self.opt.entity.base_consumptions:
                                    base_ref_amount = self.opt.entity.base_consumptions.get(ref_res, 0.0) * process.consumption_shares.get(ref_res, 0.0)
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

                                if imp['type'] == 'variation' or imp['type'] == 'up':
                                    reduction = -imp['value'] * initial_val
                                elif imp['type'] == 'new':
                                    reduction = -imp['value'] * base_ref
                                else:
                                    reduction = 0
                                    
                                # Scale reduction by active units proportion
                                reduction = reduction * (act_var / process.nb_units)
                                
                                # Apply UP compounding to the reduction if UP exists for this process
                                up_rate = 0.0
                                if 'UP' in process.valid_technologies:
                                    up_tech = self.data.technologies['UP']
                                    up_imp = up_tech.impacts.get('ALL') or up_tech.impacts.get('CO2_EM')
                                    if up_imp and (up_imp['type'] == 'variation' or up_imp['type'] == 'up'):
                                        up_rate = abs(up_imp['value'])
                                reduction *= ((1.0 - up_rate) ** yr_idx)
                                
                                captured_kt += max(0.0, reduction / 1000.0)

            really_avoided_kt = max(0.0, avoided_total_kt - captured_kt)
            
            # DAC and Credits
            dac_cap = 0.0
            dac_limit_kt = 0.0
            if self.data.dac_params.active:
                dac_v = getattr(self.opt.dac_captured_vars.get(t), 'varValue', 0.0) or 0.0
                dac_cap = dac_v / 1000.0
                if self.data.dac_params.max_volume_pct < 1.0:
                    # Consistent with optimizer logic
                    ref_yr = self.data.dac_params.ref_year
                    # Prioritize historical/external reference from Entities sheet (nested dict)
                    yr_map = self.opt.entity.ref_baselines.get('CO2_EM', {})
                    if ref_yr in yr_map:
                        ref_emis = yr_map[ref_yr]
                    elif yr_map:
                        ref_emis = next(iter(yr_map.values()))
                    else:
                        ref_emis = self.opt.entity.base_emissions

                    # For simulated years, look in emis_data
                    for d in emis_data:
                        if d['Year'] == ref_yr:
                            ref_emis = d['Direct_CO2']
                            break
                    dac_limit_kt = (ref_emis * self.data.dac_params.max_volume_pct) / 1000.0

            credit_vol = 0.0
            if self.data.credit_params.active:
                cred_v = getattr(self.opt.credit_purchased_vars.get(t), 'varValue', 0.0) or 0.0
                credit_vol = cred_v / 1000.0
                
            emis_data.append({
                'Year': t,
                'Direct_CO2': direct_co2,
                'Indirect_CO2': indirect_co2,
                'Total_CO2': total_co2,
                'DAC_Captured_kt': dac_cap,
                'DAC_Limit_kt': dac_limit_kt,
                'Credits_Purchased_kt': credit_vol,
                'Taxed_CO2': taxed_co2,
                'Free_Quota': direct_co2 * fq_pct if fq_pct <= 1.0 else fq_pct,  # Based on actual yearly emissions
                'Tax_Price': tax_price,
                'Tax_Cost_MEuros': tax_cost_meuros,
                'Avoided_Direct_CO2_kt': avoided_total_kt,
                'Avoided_Total_CO2_kt': max(0.0, (self.opt.entity.base_emissions + self.opt.entity.base_emissions * 0.1 - total_co2) / 1000.0), # Heuristic
                'Captured_CO2_kt': captured_kt,
                'Really_Avoided_CO2_kt': really_avoided_kt,
                'DAC_Captured_kt': dac_cap,
                'Credits_Purchased_kt': credit_vol
            })
            
        df_emis = pd.DataFrame(emis_data)
        df_emis['Cumul_Tax_MEuros'] = df_emis['Tax_Cost_MEuros'].cumsum()

        # Generate Indirect Emissions Breakdown
        indir_breakdown = []
        for t in self.years:
            row = {'Year': t}
            for r_id in self.data.resources:
                if r_id in self.data.time_series.other_emissions_factors:
                    factor = self.data.time_series.other_emissions_factors[r_id].get(t, 0.0)
                    if factor > 0:
                        cons_val = self.opt.cons_vars.get((t, r_id))
                        if cons_val is not None:
                            val = cons_val.varValue
                            if r_id == 'EN_ELEC':
                                base_elec = self.opt.entity.base_consumptions.get('EN_ELEC', 0.0)
                                new_elec = max(0.0, val - base_elec)
                                row['EN_ELEC_BASE'] = base_elec * factor
                                row['EN_ELEC_TECH'] = new_elec * factor
                            else:
                                row[r_id] = val * factor
            indir_breakdown.append(row)
        df_indir = pd.DataFrame(indir_breakdown).fillna(0.0)
        
        # 1.5 Technology Costs (CAPEX & OPEX)
        # We use a two-pass approach to distribute aids correctly.
        aids_dist = {} # (year, col_name) -> value
        tech_capex_spent = {} # (year, p_id, t_id) -> value
        tech_co2_abatement = {} # (year, p_id, t_id) -> value in tCO2 abated
        tech_invested_procs = {} # (year, p_id, t_id) -> list of process names (for labels)
        new_investments = set() # (year, p_id, t_id) -> tracking for arrows
        
        for t in self.years:
            for t_id in self.data.technologies:
                aids_dist[(t, f"Aid_GRANT_{t_id}")] = 0.0
                aids_dist[(t, f"Aid_CCFD_{t_id}")] = 0.0
            for p_id in self.opt.entity.processes:
                for t_id in self.data.technologies:
                    tech_capex_spent[(t, p_id, t_id)] = 0.0
                    tech_co2_abatement[(t, p_id, t_id)] = 0.0
                    tech_invested_procs[(t, p_id, t_id)] = []

        for t in self.years:
            for t_id, tech in self.data.technologies.items():
                if t_id == 'UP': continue
                for p_id, process in self.opt.entity.processes.items():
                    if t_id in process.valid_technologies:
                        # CAPEX calculation
                        cap_capex = 1.0
                        if tech.capex_per_unit:
                            if tech.capex_unit == 'tCO2': cap_capex = self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                            elif 'MW' in tech.capex_unit.upper(): 
                                cap_capex = (self.opt.entity.base_consumptions.get('EN_FUEL', 0.0) * process.consumption_shares.get('EN_FUEL', 0.0)) / self.opt.entity.annual_operating_hours
                        
                        invest_v = getattr(self.opt.invest_vars[(t, p_id, t_id)], 'varValue', 0.0) or 0.0
                        if invest_v > 1e-6:
                            current_capex = tech.capex_by_year.get(t, tech.capex)
                            true_capex = (current_capex * cap_capex / process.nb_units) * invest_v
                            if true_capex > 0 and self.verbose:
                                print(f"  [yellow][Reporter][/yellow] [DEBUG] {t}: {p_id} {t_id} invest={invest_v} capex={current_capex} true_capex={true_capex}")
                            tech_capex_spent[(t, p_id, t_id)] += true_capex
                            
                            # Calculate CO2 Abatement added by this investment
                            imp = tech.impacts.get('CO2_EM')
                            if not imp:
                                co2_keys = [k for k in tech.impacts.keys() if 'CO2' in k]
                                if co2_keys: imp = tech.impacts[co2_keys[0]]
                            
                            if imp:
                                initial_val = self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                                yr_idx = list(self.opt.years).index(t)
                                
                                # UP compounding for the baseline at time of investment
                                up_rate = 0.0
                                if 'UP' in process.valid_technologies:
                                    up_tech = self.data.technologies['UP']
                                    up_imp = up_tech.impacts.get('ALL') or up_tech.impacts.get('CO2_EM')
                                    if up_imp and (up_imp['type'] == 'variation' or up_imp['type'] == 'up'):
                                        up_rate = abs(up_imp['value'])
                                initial_val *= ((1.0 - up_rate) ** yr_idx)

                                ref_res = imp.get('ref_resource')
                                if ref_res == 'CO2_EM':
                                    base_ref = initial_val
                                elif ref_res and ref_res in self.opt.entity.base_consumptions:
                                    base_ref = self.opt.entity.base_consumptions.get(ref_res, 0.0) * process.consumption_shares.get(ref_res, 0.0)
                                else:
                                    # Fallback to initial_val if ref_res not found or same as CO2
                                    base_ref = initial_val if initial_val > 0 else 1.0

                                if imp['type'] == 'variation' or imp['type'] == 'up':
                                    red = -imp['value'] * initial_val
                                elif imp['type'] == 'new':
                                    red = -imp['value'] * base_ref
                                else:
                                    red = 0
                                
                                tech_co2_abatement[(t, p_id, t_id)] += red * (invest_v / process.nb_units)
                            
                            new_investments.add((t, p_id, t_id)) # Mark as NEW for arrows
                            
                            proc_display = f"({invest_v:.2f}/{process.nb_units})" if invest_v % 1 != 0 else f"({int(invest_v)}/{process.nb_units})"
                            tech_invested_procs[(t, p_id, t_id)].append(proc_display)
                            
                            # GRANT: Use precise value from the lp variable
                            if hasattr(self.opt, 'grant_amt_vars') and self.opt.grant_amt_vars.get((t, p_id, t_id)):
                                grant_val = getattr(self.opt.grant_amt_vars[(t, p_id, t_id)], 'varValue', 0.0) or 0.0
                                aids_dist[(t, f"Aid_GRANT_{t_id}")] += grant_val
                                    
                            # CCfD: Distributed over contract duration
                            if hasattr(self.opt, 'ccfd_used_vars') and self.opt.ccfd_used_vars.get((t, p_id, t_id)):
                                if getattr(self.opt.ccfd_used_vars[(t, p_id, t_id)], 'varValue', 0) > 0.5:
                                    ccfd_p = self.data.ccfd_params
                                    imp = tech.impacts.get('CO2_EM')
                                    if not imp:
                                        co2_keys = [k for k in tech.impacts.keys() if 'CO2' in k]
                                        if co2_keys: imp = tech.impacts[co2_keys[0]]
                                        
                                    if imp and (imp['type'] == 'variation' or imp['type'] == 'up'):
                                        reduction_frac = (-imp['value'] if imp['value'] < 0 else 0) / process.nb_units
                                        if reduction_frac > 0:
                                            max_emis = self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                                            # Scale avoided tons by units invested
                                            avoided_tons = reduction_frac * max_emis * invest_v
                                            start_yr = t + tech.implementation_time
                                            end_yr = start_yr + ccfd_p.duration
                                            
                                            for tau in range(start_yr, end_yr):
                                                if tau in self.years:
                                                    c_price = self.data.time_series.carbon_prices.get(tau, 0.0)
                                                    strike_price = (1.0 + ccfd_p.eua_price_pct) * self.data.time_series.carbon_prices.get(t, 0.0)
                                                    if ccfd_p.contract_type == 2:
                                                        subsidy = strike_price - c_price
                                                    else:
                                                        subsidy = max(0, strike_price - c_price)
                                                    aids_dist[(tau, f"Aid_CCFD_{t_id}")] += (subsidy * avoided_tons)

            if self.data.dac_params.active:
                dac_added_v = getattr(self.opt.dac_added_capacity_vars.get(t), 'varValue', 0.0) or 0.0
                if dac_added_v > 1e-4:
                    dac_capex = dac_added_v * self.data.dac_params.capex_by_year.get(t, 0.0)
                    tech_capex_spent[(t, 'INDIRECT', 'DAC')] = dac_capex
                    tech_co2_abatement[(t, 'INDIRECT', 'DAC')] = dac_added_v
                    new_investments.add((t, 'INDIRECT', 'DAC'))
                    tech_invested_procs[(t, 'INDIRECT', 'DAC')] = [f"{dac_added_v / 1000:,.0f}ktCO2"]
            
            if self.data.credit_params.active:
                # Credits are purely OPEX, but if there's a fixed setup cost (not currently in model), we'd put it here.
                # For now, ensure it's in project_ids if used.
                cred_v = getattr(self.opt.credit_purchased_vars.get(t), 'varValue', 0.0) or 0.0
                if cred_v > 1e-4:
                    tech_capex_spent[(t, 'INDIRECT', 'CREDIT')] = 0.0 # No CAPEX for credits
                    tech_invested_procs[(t, 'INDIRECT', 'CREDIT')] = [f"{cred_v / 1000:,.0f}ktCO2"]

        # Identify all relevant (p_id, t_id) projects for granular dataset
        project_ids = sorted(list({(p_id, t_id) for (y, p_id, t_id), v in tech_capex_spent.items() if v > 1e-9}))
        if self.verbose:
            print(f"  [yellow][Reporter][/yellow] [DEBUG] Found {len(project_ids)} projects with non-zero investment.")
            if len(project_ids) > 0:
                print(f"  [yellow][Reporter][/yellow] [DEBUG] Project IDs: {project_ids[:5]}...")
        
        cost_data = []      # For Excel (Aggregated by Tech)
        project_data = []   # For Plotting (Granular Process + Tech)
        
        for t in self.years:
            row_agg = {'Year': t}
            row_proj = {'Year': t}
            
            # Aggregated row (Excel)
            for t_id in self.data.technologies:
                row_agg[t_id] = sum(tech_capex_spent.get((t, p_id, t_id), 0.0) for p_id in self.opt.entity.processes)
                # Group labels by tech
                all_labels = []
                for p_id in self.opt.entity.processes:
                    all_labels.extend(tech_invested_procs.get((t, p_id, t_id), []))
                row_agg[f"{t_id}_labels"] = ", ".join(sorted(list(set(all_labels))))
                row_agg[f"Aid_GRANT_{t_id}"] = aids_dist.get((t, f"Aid_GRANT_{t_id}"), 0.0)
                row_agg[f"Aid_CCFD_{t_id}"] = aids_dist.get((t, f"Aid_CCFD_{t_id}"), 0.0)
            
            if self.data.dac_params.active:
                row_agg['DAC'] = tech_capex_spent.get((t, 'INDIRECT', 'DAC'), 0.0)
                row_agg['DAC_labels'] = ", ".join(tech_invested_procs.get((t, 'INDIRECT', 'DAC'), []))
                dac_total_v = getattr(self.opt.dac_total_capacity_vars.get(t), 'varValue', 0.0) or 0.0
                dac_captured_v = getattr(self.opt.dac_captured_vars.get(t), 'varValue', 0.0) or 0.0
                base_opex = dac_total_v * self.data.dac_params.opex_by_year.get(t, 0.0)
                elec_cons = dac_captured_v * self.data.dac_params.elec_by_year.get(t, 0.0)
                elec_price = self.data.time_series.resource_prices.get('EN_ELEC', {}).get(t, 0.0)
                row_agg['DAC_Opex'] = base_opex + (elec_cons * elec_price)
            if self.data.credit_params.active:
                cred_v = getattr(self.opt.credit_purchased_vars.get(t), 'varValue', 0.0) or 0.0
                row_agg['Credit_Cost'] = cred_v * self.data.credit_params.cost_by_year.get(t, 0.0)
            
            # Granular row (Plotting)
            for p_id, t_id in project_ids:
                col_name = f"{p_id}##{t_id}"
                row_proj[col_name] = tech_capex_spent.get((t, p_id, t_id), 0.0)
                row_proj[f"{col_name}##tCO2"] = tech_co2_abatement.get((t, p_id, t_id), 0.0)
                row_proj[f"{col_name}_is_new"] = 1 if (t, p_id, t_id) in new_investments else 0
                row_proj[f"{col_name}_labels"] = ", ".join(tech_invested_procs.get((t, p_id, t_id), []))

            cost_data.append(row_agg)
            project_data.append(row_proj)

        self.df_costs = pd.DataFrame(cost_data)
        self.df_projects = pd.DataFrame(project_data)
        df_costs = self.df_costs
        for t in self.years:
            # Compute Annual Budget Limit directly: CA revenue × ca_percentage_limit
            # This is independent of whether the solver created CA/Budget variables
            budget_val = 0.0
            ca_pct = self.opt.entity.ca_percentage_limit
            if t == self.years[0] and self.verbose:
                print(f"  [yellow][Reporter][/yellow] [DEBUG] Entity '{self.opt.entity.id}' budget limit factor: {ca_pct*100:.4f}%")
            if ca_pct > 0:
                # Try solver variable first
                ca_var_name = f"CA_{t}"
                yearly_var_name = f"YearlyBudget_{t}"
                vars_dict = self.opt.model.variablesDict()
                if ca_var_name in vars_dict and vars_dict[ca_var_name].varValue is not None:
                    budget_val = abs(vars_dict[ca_var_name].varValue) * ca_pct
                elif yearly_var_name in vars_dict and vars_dict[yearly_var_name].varValue is not None:
                    budget_val = vars_dict[yearly_var_name].varValue
                else:
                    # Fallback
                    for res_id in self.opt.entity.sold_resources:
                        price = self.data.time_series.resource_prices.get(res_id, {}).get(t, 0.0)
                        base_cons = abs(self.opt.entity.base_consumptions.get(res_id, 0.0))
                        budget_val += price * base_cons * ca_pct
            
            self.df_costs.loc[self.df_costs['Year'] == t, 'Budget_Limit'] = budget_val
            # Add Total Investment Limit (Budget + Loans, typically 1.5x Budget per optimizer logic)
            self.df_costs.loc[self.df_costs['Year'] == t, 'Total_Limit'] = 1.5 * budget_val
        
        # 1.6 Financing Data (Detailed Breakdown)
        financing_data = []
        annual_principal = {t: 0.0 for t in self.years}
        annual_interest = {t: 0.0 for t in self.years}
        # Track principal repayment per granular project to distribute back in the chart
        # Include 'INDIRECT' for DAC/Credit projects
        tech_principal_repayment = {(t, p_id, t_id): 0.0 for t in self.years for p_id in list(self.opt.entity.processes.keys()) + ['INDIRECT'] for t_id in list(self.data.technologies.keys()) + ['DAC', 'CREDIT']}
        
        loan_taken_per_year = {t: sum(getattr(self.opt.loan_vars.get((t, p_id, t_id, l_id)), 'varValue', 0.0) or 0.0 
                                      for p_id, p in self.opt.entity.processes.items() 
                                      for t_id in p.valid_technologies if t_id != 'UP'
                                      for l_id in range(len(self.data.bank_loans))) for t in self.years}

        # First pass to calculate annual totals and technology shares for distribution
        for tau in self.years:
            # Calculate project shares for this year's investment
            project_cols = [f"{p_id}##{t_id}" for p_id, t_id in project_ids]
            total_capex_tau = self.df_projects.loc[self.df_projects['Year'] == tau, project_cols].sum(axis=1).values[0]
            project_shares = {}
            if total_capex_tau > 1e-4:
                for p_id, t_id in project_ids:
                    col_name = f"{p_id}##{t_id}"
                    project_shares[(p_id, t_id)] = self.df_projects.loc[self.df_projects['Year'] == tau, col_name].values[0] / total_capex_tau

            for l_id, loan in enumerate(self.data.bank_loans):
                p_initial = sum(getattr(self.opt.loan_vars.get((tau, p_id, t_id, l_id)), 'varValue', 0.0) or 0.0
                                for p_id, p in self.opt.entity.processes.items()
                                for t_id in p.valid_technologies if t_id != 'UP')
                if p_initial > 0:
                    r = loan.rate
                    d = loan.duration
                    if r > 0:
                        annuity = p_initial * (r / (1 - (1 + r)**(-d)))
                    else:
                        annuity = p_initial / d
                    
                    current_balance = p_initial
                    for k in range(d):
                        t_rep = tau + k
                        if t_rep in self.years:
                            interest = current_balance * r
                            principal = annuity - interest
                            annual_principal[t_rep] += principal
                            annual_interest[t_rep] += interest
                            
                            # Distribute principal back to projects based on shares at time tau
                            for (p_id, t_id), share in project_shares.items():
                                tech_principal_repayment[(t_rep, p_id, t_id)] += principal * share
                                
                            current_balance -= principal

        # Construct financing dataframe and adjust cost data
        for t in self.years:
            tech_cols = [t_id for t_id in self.data.technologies if t_id != 'UP' and t_id in df_costs.columns]
            if self.data.dac_params.active and 'DAC' in df_costs.columns:
                tech_cols.append('DAC')
            total_capex = df_costs.loc[df_costs['Year'] == t, tech_cols].sum(axis=1).values[0]
            loan_taken = loan_taken_per_year[t]
            
            financing_data.append({
                'Year': t,
                'Total_CAPEX (M€)': total_capex / 1_000_000.0,
                'Loan_Principal_Taken (M€)': loan_taken / 1_000_000.0,
                'Out_of_Pocket_CAPEX (M€)': (total_capex - loan_taken) / 1_000_000.0,
                'Principal_Repayment (M€)': annual_principal[t] / 1_000_000.0,
                'Interest_Paid (M€)': annual_interest[t] / 1_000_000.0,
                'Total_Annuity (M€)': (annual_principal[t] + annual_interest[t]) / 1_000_000.0
            })
            
            # Adjust technology columns: Out-of-Pocket (current year) + Principal Repayment (from previous loans)
            for t_id in tech_cols:
                # 1. Current year out-of-pocket portion
                    if total_capex > 1e-4:
                        ratio_oop = (total_capex - loan_taken) / total_capex
                        current_oop = df_costs.loc[df_costs['Year'] == t, t_id].values[0] * ratio_oop
                    else:
                        current_oop = 0.0
                        
                    # 2. Add repayments for this technology due this year
                    repayment_portion = sum(tech_principal_repayment.get((t, p_id, t_id), 0.0) for p_id in self.opt.entity.processes)
                    
                    df_costs.loc[df_costs['Year'] == t, t_id] = current_oop + repayment_portion

            # The user requested that the Investment Plan chart only shows raw investment,
            # ignoring loan smoothing. So we do NOT modify self.df_projects here anymore.
            # self.df_projects retains the original `tech_capex_spent` raw values.
            # Add interest to df_costs for the stacked bar chart
            df_costs.loc[df_costs['Year'] == t, 'Financing Interests'] = annual_interest[t]
            # Also add to df_projects for consistency with the plotter's expectations
            self.df_projects.loc[self.df_projects['Year'] == t, 'Financing Interests'] = annual_interest[t]

        df_finance = pd.DataFrame(financing_data)
        
        # Aggregate CCfD Revenue per year for the Carbon Tax Graph
        ccfd_annual_revenue = {}
        for t in self.years:
            ccfd_annual_revenue[t] = sum(aids_dist.get((t, f"Aid_CCFD_{t_id}"), 0.0) for t_id in self.data.technologies) / 1_000_000.0
            
        df_emis['CCfD_Refund_MEuros'] = df_emis['Year'].map(ccfd_annual_revenue)
        df_emis['Net_Tax_Cost_MEuros'] = df_emis['Tax_Cost_MEuros'] - df_emis['CCfD_Refund_MEuros']
        
        # 1.6 Check for Missed Objectives (Soft Constraints)
        if self.verbose:
            print("\n  [yellow][Reporter][/yellow] [bold underline]--- Goal Achievement Report ---[/bold underline]")
        missed_goals = False
        for i, obj in enumerate(self.data.objectives):
            if obj.mode == 'LINEAR':
                t_target = min(obj.target_year, self.opt.years[-1])
                penalty_val = getattr(self.opt.penalty_vars.get((i, t_target)), 'varValue', 0.0) or 0.0
                if penalty_val > 1e-4:
                    missed_goals = True
                    if self.verbose:
                        print(f"  [bold red][!] WARNING: Objective {i+1} MISSED at year {t_target}![/bold red]")
                        print(f"    Target End: [cyan]{obj.target_year}[/cyan] | Resource: [cyan]{obj.resource}[/cyan] | Limit: [cyan]{obj.cap_value}[/cyan]")
                        print(f"    Shortfall (Penalty Paid): [bold red]{penalty_val:,.2f}[/bold red] {self.data.resources.get(obj.resource).unit if obj.resource in self.data.resources else 'units'}")
            else:
                penalty_val = getattr(self.opt.penalty_vars.get(i), 'varValue', 0.0) or 0.0
                if penalty_val > 1e-4:
                    missed_goals = True
                    if self.verbose:
                        print(f"  [bold red][!] WARNING: Objective {i+1} MISSED![/bold red]")
                        print(f"    Target: [cyan]{obj.target_year}[/cyan] | Resource: [cyan]{obj.resource}[/cyan] | Limit: [cyan]{obj.cap_value}[/cyan]")
                        print(f"    Shortfall (Penalty Paid): [bold red]{penalty_val:,.2f}[/bold red] {self.data.resources.get(obj.resource).unit if obj.resource in self.data.resources else 'units'}")
        
        if not missed_goals:
            if self.verbose:
                print("  [bold green][OK] SUCCESS: All stated objectives were successfully met by the model.[/bold green]")
        if self.verbose:
            print("  [yellow][Reporter][/yellow] -------------------------------\n")
            
            # --- NEW: H2 Sourcing Breakdown Summary in Console ---
            h2_buy_cols = [c for c in df_cons.columns if 'H2' in c.upper() and c not in ['Year', 'EN_H2_C', 'EN_H2_P', 'EN_ELEC_FOR_H2']]
            if h2_buy_cols and (df_cons[h2_buy_cols] > 1.0).any().any():
                print("\n  >>> HYDROGEN SOURCING BREAKDOWN (Optimal Mix) <<<")
                for t in self.years:
                    row = df_cons[df_cons['Year'] == t]
                    total_h2 = row[h2_buy_cols].sum(axis=1).values[0]
                    if total_h2 > 1.0:
                        shares = []
                        for col in h2_buy_cols:
                            val = row[col].values[0]
                            if val > 1.0:
                                display_name = col.replace('EN_', '').replace('_C', '').replace('_', ' ')
                                if display_name == 'H2 ON SITE': display_name = 'ON-SITE ELECTROLYSIS'
                                shares.append(f"[cyan]{display_name}[/cyan]: {val:,.0f} units ({val/total_h2*100:.1f}%)")
                        if shares:
                            print(f"  Year {t}: {', '.join(shares)}")
                    print("")
        
        
        # 1.5 Prepare Data Used sheet
        data_used_rows = []
        # Get all unique resource IDs
        all_resources = set(self.data.time_series.resource_prices.keys()).union(set(self.data.time_series.other_emissions_factors.keys()))
        for r_id in all_resources:
            for t in self.years:
                price = self.data.time_series.resource_prices.get(r_id, {}).get(t, 0.0)
                emissions = self.data.time_series.other_emissions_factors.get(r_id, {}).get(t, 0.0)
                data_used_rows.append({
                    "Resource": r_id,
                    "Year": t,
                    "Price": price,
                    "CO2_Emissions": emissions
                })
        df_data_used = pd.DataFrame(data_used_rows) if data_used_rows else pd.DataFrame(columns=["Resource", "Year", "Price", "CO2_Emissions"])

        # 2. Export to Excel
        excel_path = os.path.join(self.results_dir, 'Master_Plan.xlsx')
        if self.generate_excel or (hasattr(self.data, 'reporting_toggles') and self.data.reporting_toggles.results_excel):
            try:
                with pd.ExcelWriter(excel_path) as writer:
                    if not df_invest.empty:
                        df_invest.to_excel(writer, sheet_name='Investments', index=False)
                    df_cons.to_excel(writer, sheet_name='Energy_Mix', index=False)
                    df_emis.to_excel(writer, sheet_name='CO2_Trajectory', index=False)
                    if not df_indir.empty:
                        df_indir.to_excel(writer, sheet_name='Indirect_Emissions', index=False)
                    df_costs.to_excel(writer, sheet_name='Technology_Costs', index=False)
                    df_finance.to_excel(writer, sheet_name='Financing', index=False)
                    
                    if not df_data_used.empty:
                        df_data_used.to_excel(writer, sheet_name='Data_Used', index=False)
                    
                    # 2.2 Export Chart Data
                    if hasattr(self, 'charts_data') and self.charts_data:
                        start_row = 0
                        sheet_name = 'Charts'
                        for title, df_chart in self.charts_data:
                            # Write Title
                            pd.Series([title]).to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False, header=False)
                            # Write Data
                            df_chart.to_excel(writer, sheet_name=sheet_name, startrow=start_row + 1, index=True)
                            # Advance 2 blank lines (title + header + data rows + 2)
                            start_row += len(df_chart) + 4
                    
                if self.verbose:
                    print(f"  [yellow][Reporter][/yellow] [OK] Exported data to [bold cyan]{excel_path}[/bold cyan]")
            except PermissionError:
                os.makedirs(results_dir, exist_ok=True)
                excel_path = os.path.join(results_dir, 'Master_Plan_new.xlsx')
                with pd.ExcelWriter(excel_path) as writer:
                    if not df_invest.empty:
                        df_invest.to_excel(writer, sheet_name='Investments', index=False)
                    df_cons.to_excel(writer, sheet_name='Energy_Mix', index=False)
                    df_emis.to_excel(writer, sheet_name='CO2_Trajectory', index=False)
                    if not df_indir.empty:
                        df_indir.to_excel(writer, sheet_name='Indirect_Emissions', index=False)
                    df_costs.to_excel(writer, sheet_name='Technology_Costs', index=False)
                    df_finance.to_excel(writer, sheet_name='Financing', index=False)
    
                    if not df_data_used.empty:
                        df_data_used.to_excel(writer, sheet_name='Data_Used', index=False)
                        
                    # 2.2 Export Chart Data (Fallback)
                    if hasattr(self, 'charts_data') and self.charts_data:
                        start_row = 0
                        sheet_name = 'Charts'
                        for title, df_chart in self.charts_data:
                            pd.Series([title]).to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False, header=False)
                            df_chart.to_excel(writer, sheet_name=sheet_name, startrow=start_row + 1, index=True)
                            start_row += len(df_chart) + 4
    
                if self.verbose:
                    print(f"  [yellow][Reporter][/yellow] [!] File was locked. Exported data to [bold cyan]{excel_path}[/bold cyan] instead.")
        # 3. NEW: CO2 Abatement Cost Calculation (MAC) - Aggregate by Tech Type
        # ── 1. HELPERS: Uniform effective price and emission factor ──────────────────────
        # ALL resources get the same treatment:
        #   effective_price = market_price + emission_factor × carbon_tax  (Scope 3 included)
        #   effective_emiss = emission_factor from time series
        # No hardcoded resource IDs — the logic reads whatever is in the data.

        def _get_effective_price(r_id, yr):
            """All-in cost for one unit of resource r_id at year yr.
            Includes the market price plus the indirect carbon cost of upstream emissions."""
            market_p = self.data.time_series.resource_prices.get(r_id, {}).get(yr, 0.0)
            emiss_f  = self.data.time_series.other_emissions_factors.get(r_id, {}).get(yr, 0.0)
            carbon_p = self.data.time_series.carbon_prices.get(yr, 0.0)
            return market_p + (emiss_f * carbon_p)

        def _get_effective_emiss(r_id, yr):
            """Indirect emission factor (tCO2/unit) for resource r_id at year yr."""
            return self.data.time_series.other_emissions_factors.get(r_id, {}).get(yr, 0.0)

        def _get_metric(r_id, yr, metric_type):
            """Unified metric getter — same logic for every resource."""
            if metric_type == 'price':
                return _get_effective_price(r_id, yr)
            else:
                return _get_effective_emiss(r_id, yr)

        def _primary_energy_consumption(process):
            """Return the base consumption of the dominant energy resource for a process.
            Searches all base_consumptions ranked by share — avoids hardcoding 'EN_FUEL'."""
            best_res, best_val = None, 0.0
            for res, share in process.consumption_shares.items():
                base = self.opt.entity.base_consumptions.get(res, 0.0)
                val  = base * share
                if val > best_val:
                    best_val = val
                    best_res = res
            return best_val, best_res

        # ── 2. AGGREGATE MAC BY TECHNOLOGY ───────────────────────────────────────────────
        tech_mac_agg = {}   # t_id → {capex, opex, co2, processes, status}

        all_techs = [t_id for t_id in self.data.technologies.keys() if t_id != 'UP']
        if self.data.dac_params.active:    all_techs.append('DAC')
        if self.data.credit_params.active: all_techs.append('CREDIT')

        for t_id in all_techs:
            if t_id not in tech_mac_agg:
                tech_mac_agg[t_id] = {'capex': 0.0, 'opex': 0.0, 'co2': 0.0,
                                      'processes': [], 'status': 'Potential'}

            if   t_id == 'DAC':    valid_p_ids = ['INDIRECT']
            elif t_id == 'CREDIT': valid_p_ids = ['INDIRECT']
            else:
                valid_p_ids = [p_id for p_id, p in self.opt.entity.processes.items()
                               if t_id in p.valid_technologies]

            is_invested = any((p_id, t_id) in project_ids for p_id in valid_p_ids)

            if is_invested:
                # ── INVESTED path ────────────────────────────────────────────
                tech_mac_agg[t_id]['status'] = 'Invested'
                tech = self.data.technologies.get(t_id)

                for p_id in valid_p_ids:
                    if (p_id, t_id) not in project_ids:
                        continue
                    process = self.opt.entity.processes.get(p_id)
                    if not tech or not process:
                        continue

                    # A — CAPEX (actual spent)
                    tech_mac_agg[t_id]['capex'] += sum(
                        tech_capex_spent.get((t, p_id, t_id), 0.0) for t in self.years)

                    # B — CO2 Abated (uses actual activation variables)
                    for res_id, imp in tech.impacts.items():
                        for t in self.years:
                            act_var = getattr(self.opt.active_vars.get((t, p_id, t_id)),
                                              'varValue', 0.0) or 0.0
                            if act_var < 1e-6:
                                continue

                            if res_id == 'CO2_EM' or 'CO2' in res_id.upper():
                                y_factor = 1.0
                            else:
                                y_factor = _get_effective_emiss(res_id, t)
                                if y_factor == 0.0:
                                    continue

                            if res_id == 'CO2_EM' or 'CO2' in res_id.upper():
                                base_val = (self.opt.entity.base_emissions
                                            * process.emission_shares.get('CO2_EM', 0.0))
                            else:
                                base_val = (self.opt.entity.base_consumptions.get(res_id, 0.0)
                                            * process.consumption_shares.get(res_id, 0.0))

                            yr_idx  = list(self.opt.years).index(t)
                            up_rate = 0.0
                            if 'UP' in process.valid_technologies:
                                up_tech = self.data.technologies['UP']
                                up_imp  = up_tech.impacts.get('ALL') or up_tech.impacts.get(res_id)
                                if up_imp and up_imp['type'] in ['variation', 'up']:
                                    up_rate = abs(up_imp['value'])
                            base_val *= (1.0 - up_rate) ** yr_idx

                            if imp['type'] in ['variation', 'up']:
                                target_red = -imp['value'] * base_val
                            elif imp['type'] == 'new':
                                ref_res  = imp.get('ref_resource')
                                base_ref = (self.opt.entity.base_consumptions.get(ref_res, 0.0)
                                            * process.consumption_shares.get(ref_res, 0.0)
                                            if ref_res and ref_res in self.opt.entity.base_consumptions
                                            else base_val)
                                target_red = -imp['value'] * base_ref
                            else:
                                target_red = 0.0

                            tech_mac_agg[t_id]['co2'] += (target_red * y_factor) * (act_var / process.nb_units)

                    # C — OPEX Change (actual, summed over all active years)
                    for t in self.years:
                        act_var = getattr(self.opt.active_vars.get((t, p_id, t_id)),
                                          'varValue', 0.0) or 0.0
                        if act_var < 1e-6:
                            continue

                        year_opex_change = 0.0
                        for res_id, imp_op in tech.impacts.items():
                            if res_id in ('CO2_EM', 'ALL'):
                                continue

                            ref_res   = imp_op.get('ref_resource') or res_id
                            ref_state = imp_op.get('reference', 'INITIAL')

                            if ref_res == 'CO2_EM':
                                base_ref = (self.opt.entity.base_emissions
                                            * process.emission_shares.get('CO2_EM', 0.0))
                            else:
                                base_ref = (self.opt.entity.base_consumptions.get(ref_res, 0.0)
                                            * process.consumption_shares.get(ref_res, 0.0))

                            if imp_op['type'] in ('variation', 'up'):
                                target_change = imp_op['value'] * base_ref
                            elif imp_op['type'] == 'new':
                                if ref_state == 'AVOIDED':
                                    co2_imp = tech.impacts.get('CO2_EM')
                                    if co2_imp and co2_imp['type'] == 'variation' and (ref_res == 'CO2_EM' or not ref_res):
                                        target_change = imp_op['value'] * abs(co2_imp['value'] * base_ref)
                                    else:
                                        target_change = imp_op['value'] * base_ref
                                else:
                                    target_change = imp_op['value'] * base_ref
                            else:
                                target_change = 0.0

                            year_opex_change += (target_change * (act_var / process.nb_units)) \
                                                * _get_effective_price(res_id, t)

                        current_opex = tech.opex_by_year.get(t, tech.opex)
                        cap_opex     = 1.0
                        if tech.opex_per_unit:
                            if tech.opex_unit == 'tCO2':
                                cap_opex = (self.opt.entity.base_emissions
                                            * process.emission_shares.get('CO2_EM', 0.0))
                            elif 'MW' in tech.opex_unit.upper():
                                primary_conso, _ = _primary_energy_consumption(process)
                                cap_opex = primary_conso / self.opt.entity.annual_operating_hours

                        tech_mac_agg[t_id]['opex'] += (
                            year_opex_change + (current_opex * cap_opex / process.nb_units) * act_var)

                    tech_mac_agg[t_id]['processes'].append(p_id)

            else:
                # ── POTENTIAL path ───────────────────────────────────────────
                t_eval = self.years[0]
                tech   = self.data.technologies.get(t_id)
                if not tech:
                    continue

                for p_id in valid_p_ids:
                    process = self.opt.entity.processes.get(p_id)
                    if not process:
                        continue

                    # A — Potential CAPEX
                    cap_capex = 1.0
                    if tech.capex_per_unit:
                        if tech.capex_unit == 'tCO2':
                            cap_capex = (self.opt.entity.base_emissions
                                         * process.emission_shares.get('CO2_EM', 0.0))
                        elif 'MW' in tech.capex_unit.upper():
                            primary_conso, _ = _primary_energy_consumption(process)
                            cap_capex = primary_conso / self.opt.entity.annual_operating_hours
                    tech_mac_agg[t_id]['capex'] += tech.capex_by_year.get(t_eval, tech.capex) * cap_capex

                    # B — Potential CO2 Abated
                    for res_id, imp in tech.impacts.items():
                        if res_id == 'CO2_EM' or 'CO2' in res_id.upper():
                            factor = 1.0
                        else:
                            factor = _get_effective_emiss(res_id, t_eval)
                            if factor == 0.0:
                                continue

                        if res_id == 'CO2_EM' or 'CO2' in res_id.upper():
                            base_val = (self.opt.entity.base_emissions
                                        * process.emission_shares.get('CO2_EM', 0.0))
                        else:
                            base_val = (self.opt.entity.base_consumptions.get(res_id, 0.0)
                                        * process.consumption_shares.get(res_id, 0.0))

                        if imp['type'] in ('variation', 'up'):
                            target_red = -imp['value'] * base_val
                        elif imp['type'] == 'new':
                            ref_res  = imp.get('ref_resource')
                            base_ref = (self.opt.entity.base_consumptions.get(ref_res, 0.0)
                                        * process.consumption_shares.get(ref_res, 0.0)
                                        if ref_res and ref_res in self.opt.entity.base_consumptions
                                        else base_val)
                            target_red = -imp['value'] * base_ref
                        else:
                            target_red = 0.0

                        tech_mac_agg[t_id]['co2'] += (target_red * factor) \
                                                     * (len(self.years) - tech.implementation_time)

                    # C — Potential OPEX Change
                    for t in self.years[tech.implementation_time:]:
                        year_opex_change = 0.0
                        for res_id, imp_op in tech.impacts.items():
                            if res_id in ('CO2_EM', 'ALL'):
                                continue

                            ref_res   = imp_op.get('ref_resource') or res_id
                            ref_state = imp_op.get('reference', 'INITIAL')

                            if ref_res == 'CO2_EM':
                                base_ref = (self.opt.entity.base_emissions
                                            * process.emission_shares.get('CO2_EM', 0.0))
                            else:
                                base_ref = (self.opt.entity.base_consumptions.get(ref_res, 0.0)
                                            * process.consumption_shares.get(ref_res, 0.0))

                            if imp_op['type'] in ('variation', 'up'):
                                target_change = imp_op['value'] * base_ref
                            elif imp_op['type'] == 'new':
                                if ref_state == 'AVOIDED':
                                    co2_imp = tech.impacts.get('CO2_EM')
                                    if co2_imp and co2_imp['type'] == 'variation' and (ref_res == 'CO2_EM' or not ref_res):
                                        target_change = imp_op['value'] * abs(co2_imp['value'] * base_ref)
                                    else:
                                        target_change = imp_op['value'] * base_ref
                                else:
                                    target_change = imp_op['value'] * base_ref
                            else:
                                target_change = 0.0

                            year_opex_change += target_change * _get_effective_price(res_id, t)

                        cur_opex = tech.opex_by_year.get(t, tech.opex)
                        cap_opex = 1.0
                        if tech.opex_per_unit:
                            if tech.opex_unit == 'tCO2':
                                cap_opex = (self.opt.entity.base_emissions
                                            * process.emission_shares.get('CO2_EM', 0.0))
                            elif 'MW' in tech.opex_unit.upper():
                                primary_conso, _ = _primary_energy_consumption(process)
                                cap_opex = primary_conso / self.opt.entity.annual_operating_hours

                        tech_mac_agg[t_id]['opex'] += year_opex_change + (cur_opex * cap_opex)

                    tech_mac_agg[t_id]['processes'].append(p_id)

        # DAC and CREDIT special potential logic
        if 'DAC' in tech_mac_agg and tech_mac_agg['DAC']['status'] == 'Potential':
            tech_mac_agg['DAC']['co2']   = 100_000.0 * (len(self.years) - 2)
            tech_mac_agg['DAC']['capex'] = 100_000.0 * self.data.dac_params.capex_by_year.get(self.years[0], 500.0)
            tech_mac_agg['DAC']['opex']  = 100_000.0 * self.data.dac_params.opex_by_year.get(self.years[0], 100.0) * (len(self.years) - 2)
        if 'CREDIT' in tech_mac_agg and tech_mac_agg['CREDIT']['status'] == 'Potential':
            tech_mac_agg['CREDIT']['co2']  = 50_000.0 * len(self.years)
            tech_mac_agg['CREDIT']['opex'] = 50_000.0 * self.data.credit_params.cost_by_year.get(self.years[0], 50.0) * len(self.years)

        # ── 3. BUILD MAC DATA LIST ────────────────────────────────────────────────────────
        # One entry per technology ID — no decomposition by sub-resource.
        # Each FUEL_TO_H2_X variant is already its own t_id and produces its own bar.
        mac_data = []
        for t_id, vals in tech_mac_agg.items():
            if vals['co2'] < 1.0:
                continue

            if t_id in self.data.technologies:
                tech_name = self.data.technologies[t_id].name
            elif t_id == 'DAC':
                tech_name = "Direct Air Capture"
            elif t_id == 'CREDIT':
                tech_name = "Carbon Credits"
            else:
                tech_name = t_id

            proc_sum = ", ".join(sorted(set(vals['processes']))[:3])
            if len(set(vals['processes'])) > 3:
                proc_sum += ", ..."

            mac_data.append({
                'Project':               tech_name,
                'Process Summary':       proc_sum,
                'Display Label':         f"{tech_name}\n({proc_sum})" if proc_sum else tech_name,
                'MAC (€/tCO2)':          (vals['capex'] + vals['opex']) / vals['co2'],
                'MAC CAPEX (€/tCO2)':    vals['capex'] / vals['co2'],
                'MAC OPEX (€/tCO2)':     vals['opex']  / vals['co2'],
                'Total Abated (tCO2)':   vals['co2'],
                'Total CAPEX (M€)':      vals['capex'] / 1_000_000.0,
                'Total OPEX Change (M€)': vals['opex'] / 1_000_000.0,
                'Status':                vals['status'],
            })

        # Build MAC dataframe once so chart gating can safely test emptiness.
        df_mac = pd.DataFrame(mac_data)
        if not df_mac.empty:
            df_mac = df_mac.sort_values(by='MAC (€/tCO2)').reset_index(drop=True)

        # 4. Generate Visualizations
        if self.verbose:
            print("  [yellow][Reporter][/yellow] [PLOT] Generating visualizations...")
        self.charts_data = [] # Clear/Re-init for collection
        toggles = self.data.reporting_toggles

        if toggles.chart_energy_mix: 
            self._plot_energy_mix(df_cons)
            _step()
        if toggles.chart_co2_trajectory: 
            self._plot_co2_trajectory(df_emis)
            _step()
        if toggles.chart_indirect_emissions:
            if not df_indir.empty:
                self._plot_indirect_emissions(df_indir)
            else:
                self._save_no_data_chart(
                    f'{self.scenario_name}_Indirect_Emissions.png',
                    'INDIRECT EMISSIONS',
                    'No indirect emissions data available for this scenario.'
                )
            _step()
        if toggles.chart_investment_costs: 
            self._plot_investment_costs(df_costs, df_finance)
            _step()
        if toggles.chart_total_opex: 
            self._plot_total_opex(df_cons, df_emis)
            _step()
        if toggles.chart_carbon_tax_avoided: 
            self._plot_carbon_tax_and_avoided(df_emis)
            _step()
        if toggles.chart_external_financing: 
            self._plot_external_financing(df_costs, df_finance)
            _step()
        if toggles.chart_transition_costs: 
            self._plot_transition_costs(df_costs, df_finance, df_emis)
            _step()
        if toggles.chart_carbon_prices: 
            self._plot_carbon_prices()
            _step()
        if toggles.chart_interest_paid: 
            self._plot_interest_paid(df_finance)
            _step()
        if toggles.chart_resource_prices: 
            self._plot_prices()
            _step()
        if toggles.chart_co2_abatement_cost:
            if not df_mac.empty:
                self._plot_co2_abatement_cost(df_mac)
            else:
                self._save_no_data_chart(
                    f'{self.scenario_name}_CO2_Abatement_Cost.png',
                    'MARGINAL ABATEMENT COST',
                    'No valid abatement projects found to build a MAC curve.'
                )
            _step()
        
        # 4. Final step: Re-run Excel export to include the charts sheet if needed
        # (This is a bit redundant but ensures we have the data captured AFTER plots are run)
        if self.generate_excel:
            self._export_charts_sheet(excel_path)
            _step()
        
        # Global cleanup to prevent memory leaks and Tcl threading errors
        plt.close('all')

    def _export_charts_sheet(self, excel_path: str):
        """Append or overwrite the 'Charts' sheet with collected plot data."""
        if not self.charts_data:
            return
            
        try:
            # Check if file exists to decide mode
            if os.path.exists(excel_path):
                # To avoid pandas overwrite loop bugs with replace, remove the sheet first
                from openpyxl import load_workbook
                wb = load_workbook(excel_path)
                if 'Charts' in wb.sheetnames:
                    del wb['Charts']
                    wb.save(excel_path)
                wb.close()
                with pd.ExcelWriter(excel_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                    start_row = 0
                    sheet_name = 'Charts'
                    for title, df_chart in self.charts_data:
                        if df_chart is None or df_chart.empty:
                            continue
                        pd.Series([title]).to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False, header=False)
                        df_chart.to_excel(writer, sheet_name=sheet_name, startrow=start_row + 1, index=True)
                        start_row += len(df_chart) + 4
            else:
                 with pd.ExcelWriter(excel_path) as writer:
                    start_row = 0
                    sheet_name = 'Charts'
                    for title, df_chart in self.charts_data:
                        if df_chart is None or df_chart.empty:
                            continue
                        pd.Series([title]).to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False, header=False)
                        df_chart.to_excel(writer, sheet_name=sheet_name, startrow=start_row + 1, index=True)
                        start_row += len(df_chart) + 4
        except Exception as e:
            print(f"  [yellow][Reporter][/yellow] [!] Could not export charts data: {e}")
        
    def _add_watermark(self, fig, is_dark_bg=False):
        """Add a transparent centered watermark logo to the figure."""
        try:
            logo_path = os.path.join('INPUT', 'logo_light.png' if is_dark_bg else 'logo_dark.png')
            if os.path.exists(logo_path):
                img = mpimg.imread(logo_path)
                # Add a new axes that spans the whole figure, ON TOP of other elements
                ax_logo = fig.add_axes([0, 0, 1, 1], zorder=100)
                ax_logo.axis('off')
                # Center the image with very low alpha (watermark mode)
                # extent [left, right, bottom, top] in figure coordinates (0-1)
                ax_logo.imshow(img, alpha=0.08, aspect='equal', 
                              extent=[0.2, 0.8, 0.2, 0.8], 
                              interpolation='bilinear')
        except Exception as e:
            print(f"  [red][Reporter][/red] [!] Could not add watermark: {e}")

    def _apply_premium_style(self, ax, is_dark=False):
        """Standardize the look with premium design choices."""
        # Spines
        for spine in ['top', 'right']:
            ax.spines[spine].set_visible(False)
        
        spine_color = '#EEEEEE' if is_dark else '#CCCCCC'
        for spine in ['left', 'bottom']:
            ax.spines[spine].set_color(spine_color)
            ax.spines[spine].set_linewidth(0.8)
            
        # Grid
        grid_color = '#2B2B36' if is_dark else '#E0E0E0'
        ax.grid(True, linestyle='-', alpha=0.15, color=grid_color)
        
        # Ticks
        tick_color = 'white' if is_dark else '#444444'
        ax.tick_params(colors=tick_color, labelsize=9)
        
        # Legend
        leg = ax.get_legend()
        if leg:
            if is_dark:
                leg.get_frame().set_facecolor('#0D0D14')
                leg.get_frame().set_edgecolor('#2B2B36')
                for text in leg.get_texts():
                    text.set_color('white')
            else:
                leg.get_frame().set_facecolor('white')
                leg.get_frame().set_edgecolor('#E0E0E0')
                leg.get_frame().set_alpha(0.8)
            leg.get_frame().set_linewidth(0.5)
            for text in leg.get_texts():
                text.set_fontsize(12)

    def _plot_energy_mix(self, df: pd.DataFrame):
        GJ_TO_MWH = 1.0 / 3.6  # 1 GJ = 0.2778 MWh

        df_plot = df.copy()
        df_plot.set_index('Year', inplace=True)
        # remove resources that are all zeros to declutter
        df_plot = df_plot.loc[:, (df_plot != 0).any(axis=0)]

        # ── Step 1: Convert GJ columns to MWh and merge with MWh columns ──────
        # Identify GJ and MWh columns
        gj_cols  = [c for c in df_plot.columns if c in self.data.resources and
                    str(self.data.resources[c].unit).strip().upper() == 'GJ']
        mwh_cols = [c for c in df_plot.columns if c in self.data.resources and
                    str(self.data.resources[c].unit).strip().upper() == 'MWH']

        # Convert GJ → MWh in place
        for col in gj_cols:
            df_plot[col] = df_plot[col] * GJ_TO_MWH

        # For each GJ column: if a MWh twin with the same name exists, add; otherwise just rename unit
        # We register which cols now count as MWh
        converted_to_mwh = set(gj_cols)

        # ── Step 2: Build unit grouping (GJ columns now treated as MWh) ────────
        resources_by_unit = {}
        for col in df_plot.columns:
            if col in self.data.resources:
                raw_unit = str(self.data.resources[col].unit).strip().upper()
                # Treat GJ as MWh after conversion
                unit_key = 'MWh' if raw_unit in ('GJ', 'MWH') else self.data.resources[col].unit
            else:
                unit_key = 'Unknown'

            if unit_key not in resources_by_unit:
                resources_by_unit[unit_key] = []
            resources_by_unit[unit_key].append(col)

        # ── Step 3: Rename columns to resource display names ────────────────────
        def _display_name(col: str) -> str:
            """Return the resource's human-readable name, falling back to ID."""
            if col in self.data.resources:
                res = self.data.resources[col]
                return res.name if res.name and res.name != res.id else res.id
            return col

        # Generate a plot for each unit group
        for unit, resources in resources_by_unit.items():
            df_unit = df_plot[resources].copy()

            # Rename columns to display names (unique: suffix ID in parentheses if collision)
            rename_map = {}
            seen_names = {}
            for col in df_unit.columns:
                dname = _display_name(col)
                if dname in seen_names:
                    dname = f"{dname} ({col})"
                seen_names[dname] = True
                rename_map[col] = dname
            df_unit = df_unit.rename(columns=rename_map)

            # ── Auto-scale: choose the best prefix based on the max absolute value ──
            max_val = df_unit.abs().max().max()
            if max_val >= 1_000_000:
                scale, prefix = 1_000_000, f'M {unit}'
            elif max_val >= 1_000:
                scale, prefix = 1_000, f'k {unit}'
            else:
                scale, prefix = 1, unit
            df_unit = df_unit / scale

            # Separate Consumption (positive) and Production (negative → absolute value)
            df_cons = df_unit[df_unit > 0].fillna(0)
            df_prod = df_unit[df_unit < 0].fillna(0).abs()

            # Remove entirely empty columns in both separated frames (using 1e-6 threshold against float precision bugs)
            df_cons = df_cons.loc[:, (df_cons > 1e-6).any(axis=0)]
            df_prod = df_prod.loc[:, (df_prod > 1e-6).any(axis=0)]

            # Create subplots only if there's data to plot for this unit
            has_cons = not df_cons.empty
            has_prod = not df_prod.empty

            # Strong, readable curated palettes (avoiding pale/white colors for visibility)
            cons_palette = ['#1A5276', '#2980B9', '#8E44AD', '#D35400', '#C0392B', '#273746', '#117864', '#B9770E']
            prod_palette = ['#1E8449', '#148F77', '#B7950B', '#28B463', '#239B56', '#D4AC0D', '#8C14FC', '#2C3E50']

            if has_cons:
                fig, ax1 = plt.subplots(figsize=(12, 9))
                fig.set_facecolor('white')
                
                # Use solid colors directly in pandas area plot, automatically cycling if many resources
                df_cons.plot.area(ax=ax1, stacked=True, alpha=0.85, color=cons_palette)
                
                ax1.set_title(f'CONSUMPTION ({prefix})', fontsize=13, weight='bold', pad=15)
                ax1.set_ylabel(prefix, fontsize=10, weight='semibold')
                ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:,.0f}'))
                ax1.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=True, fontsize=12)
                self._apply_premium_style(ax1)

                fig.suptitle(f'ENERGY FLOW PROFILE - {prefix}', fontsize=16, weight='bold', y=1.02)
                plt.tight_layout()
                fig.subplots_adjust(bottom=0.2)
                self._add_watermark(fig)

                os.makedirs(self.results_dir, exist_ok=True)
                safe_unit_name = str(unit).replace('/', '_').replace(' ', '_')
                plt.savefig(os.path.join(self.results_dir, f'Energy_Mix_Consumption_{safe_unit_name}.png'), dpi=300, bbox_inches='tight')
                plt.close()
                
                self.charts_data.append((f"Energy Flow Profile (Consumption) - {prefix}", df_cons))

            if has_prod:
                fig, ax2 = plt.subplots(figsize=(12, 9))
                fig.set_facecolor('white')

                df_prod.plot.area(ax=ax2, stacked=True, alpha=0.85, color=prod_palette)

                ax2.set_title(f'PRODUCTION ({prefix})', fontsize=13, weight='bold', pad=15)
                ax2.set_ylabel(prefix, fontsize=10, weight='semibold')
                ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:,.0f}'))
                ax2.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=True, fontsize=12)
                self._apply_premium_style(ax2)

                fig.suptitle(f'ENERGY FLOW PROFILE - {prefix}', fontsize=16, weight='bold', y=1.02)
                plt.tight_layout()
                fig.subplots_adjust(bottom=0.2)
                self._add_watermark(fig)

                os.makedirs(self.results_dir, exist_ok=True)
                safe_unit_name = str(unit).replace('/', '_').replace(' ', '_')
                plt.savefig(os.path.join(self.results_dir, f'Energy_Mix_Production_{safe_unit_name}.png'), dpi=300, bbox_inches='tight')
                plt.close()
                
                self.charts_data.append((f"Energy Flow Profile (Production) - {prefix}", df_prod))

        
    def _plot_co2_trajectory(self, df: pd.DataFrame):
        fig, ax1 = plt.subplots(figsize=(12, 9))
        # No tight_layout call here yet, wait for end
        
        # Convert all CO2 values to ktCO2 for readability
        KT = 1_000.0
        df = df.copy()
        for col in ['Direct_CO2', 'Indirect_CO2', 'Total_CO2', 'Taxed_CO2', 'Free_Quota']:
            if col in df.columns:
                df[col] = df[col] / KT
        
        dac_cap = df.get('DAC_Captured_kt', pd.Series(0, index=df.index))
        cred = df.get('Credits_Purchased_kt', pd.Series(0, index=df.index))
        has_dac_or_cred = (dac_cap + cred).max() > 1e-4

        # Net Direct and Net Total calculation
        net_direct = df['Direct_CO2'] - dac_cap - cred
        net_total = net_direct + df['Indirect_CO2']

        # Base lines
        ax1.plot(df['Year'], df['Direct_CO2'], label='Direct Emissions', linewidth=3, color='black', zorder=4)
        ax1.plot(df['Year'], df['Indirect_CO2'], label='Indirect Emissions', linestyle=':', linewidth=2, color='black', zorder=3)
        
        # New Total Emissions: Net Direct + Indirect
        ax1.plot(df['Year'], net_total, label='Total Emissions (Net Direct + Indirect)', linestyle='--', linewidth=3, color='darkred', zorder=4)
        
        # Net Direct Emissions
        ax1.plot(df['Year'], net_direct, label='Net Direct Emissions', linewidth=3, color='#3498db', linestyle='-.', zorder=5)
        
        # Free Quotas Surface
        ax1.fill_between(df['Year'], 0, df['Free_Quota'], alpha=0.3, color='green', label='Free Quotas (Direct)', zorder=1)
        
        # Taxed Emissions as hatched surface ON TOP of Free Quota
        ax1.fill_between(df['Year'], df['Free_Quota'], df['Free_Quota'] + df['Taxed_CO2'], 
                         hatch='...', alpha=0.4, color='tab:gray', label='Taxed Emissions (Surface)', zorder=2)
        
        # Plot Company Objectives (Goals)
        plotted_groups = set()
        
        # Color mapping for groups
        available_colors = ['#e74c3c', '#9b59b6', '#f39c12', '#1abc9c', '#34495e', '#d35400', '#2ecc71']
        group_colors = {}
        
        for obj in self.data.objectives:
            if obj.resource == 'CO2_EM':
                 # Calculate the actual limit value for plotting (convert to ktCO2)
                 if obj.comparison_year and -1.0 <= obj.cap_value <= 1.0:
                     limit = self.opt.entity.base_emissions * (1 + obj.cap_value) / KT
                 else:
                     limit = obj.cap_value / KT
                 
                 display_name = obj.name if obj.name else (obj.group if obj.group else 'Goal')
                 label = display_name if display_name not in plotted_groups else None
                 plotted_groups.add(display_name)
                 
                 # Assign dynamic color based on group
                 grp = obj.group if obj.group else 'Default'
                 if grp not in group_colors:
                     group_colors[grp] = available_colors[len(group_colors) % len(available_colors)]
                 pt_color = group_colors[grp]
                 
                 if obj.mode == 'LINEAR':
                     # The dotted trajectory line has been removed per user request
                     ax1.scatter(obj.target_year, limit, color=pt_color, marker='x', s=120, linewidths=3, label=label, zorder=6, clip_on=False)
                 else:
                     ax1.scatter(obj.target_year, limit, color=pt_color, marker='x', s=120, linewidths=3, label=label, zorder=6, clip_on=False)

        ax1.set_title('CO2 EMISSIONS TRAJECTORY & GOALS', fontsize=15, weight='bold', pad=20)
        ax1.set_ylabel('ktCO2', fontsize=12, weight='semibold')
        ax1.set_xlabel('Year', fontsize=12, weight='semibold')
        
        # Add a subtle shadow/fill under Net Direct Emissions
        ax1.fill_between(df['Year'], 0, net_direct, color='#3498db', alpha=0.1, zorder=0)

        self._apply_premium_style(ax1)
        
        if has_dac_or_cred:
            ax2 = ax1.twinx()
            
            # ── SCALE LOGIC: ax2 inverted, Origin at Top, 3x scale of max abatement ──
            max_abated = (dac_cap + cred).max()
            if max_abated < 1e-4: max_abated = 1.0
            
            # 3 times larger than the maximum of (voluntary credit + DAC Captured)
            ymax2_limit = 3 * max_abated
            
            # Origin at the top (0) and pointing downward (ymax2_limit)
            # In matplotlib, set_ylim(bottom, top) -> to invert we do (max, 0)
            ax2.set_ylim(ymax2_limit, 0)
            
            # Set ax1 limits normally (0 to max emissions + 10% margin)
            ymax1 = max(df['Direct_CO2'].max(), df['Indirect_CO2'].max(), net_direct.max(), (df['Free_Quota'] + df['Taxed_CO2']).max()) * 1.1
            if ymax1 < 1e-4: ymax1 = 1.0
            ax1.set_ylim(0, ymax1)
            
            # Use a slightly more professional palette for DAC/Credits
            dac_color = '#3498DB'
            cred_color = '#27AE60'

            # Plot filling DOWNWARDS from the top (0)
            # Since ax2 is inverted [ymax2, 0], values 0 -> ymax2 move DOWN
            ax2.fill_between(df['Year'], 0, dac_cap, color=dac_color, alpha=0.6, label='DAC Captured (ktCO2)', zorder=2)
            ax2.fill_between(df['Year'], dac_cap, dac_cap + cred, color=cred_color, alpha=0.6, label='Voluntary Credits (ktCO2)', zorder=2)
            
            ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0f}'))
            ax2.set_ylabel('DAC & CREDITS (ktCO2)', fontsize=11, color=cred_color, weight='bold')
            ax2.tick_params(axis='y', labelcolor=cred_color)
            
            # Premium styling for ax2
            for spine in ['top', 'right', 'bottom']:
                ax2.spines[spine].set_visible(False)
            ax2.spines['left'].set_visible(True)
            ax2.spines['left'].set_position(('outward', 60))
            ax2.spines['left'].set_color(cred_color)
            ax2.yaxis.set_label_position('left')
            ax2.yaxis.set_ticks_position('left')
            
            # Combine legends with premium look
            lines_1, labels_1 = ax1.get_legend_handles_labels()
            lines_2, labels_2 = ax2.get_legend_handles_labels()
            by_label = dict(zip(labels_1 + labels_2, lines_1 + lines_2))
            ax1.legend(by_label.values(), by_label.keys(), loc='upper center', bbox_to_anchor=(0.5, -0.15), 
                       ncol=min(4, len(by_label)), frameon=True, shadow=False, fontsize=12)
            
        else:
            ax1.set_ylim(bottom=0)
            ax1.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), 
                       ncol=min(4, len(by_label)) if 'by_label' in locals() else 3, frameon=True, shadow=False, fontsize=12)

        plt.tight_layout()
        fig.subplots_adjust(bottom=0.20) # Make room for the legend
        self._add_watermark(fig)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, f'{self.scenario_name}_CO2_Trajectory.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        # Store data
        self.charts_data.append(("CO2 Emissions Trajectory & Decarbonization Goals", df))

    def _plot_indirect_emissions(self, df: pd.DataFrame):
        """Plot indirect emissions breakdown by resource."""
        fig = plt.figure(figsize=(12, 9))
        # No tight_layout call here yet, wait for end
        df_plot = df.copy()
        df_plot.set_index('Year', inplace=True)
        
        # Remove resources that are all zeros
        df_plot = df_plot.loc[:, (df_plot != 0).any(axis=0)]
        
        if df_plot.empty:
            return

        # Convert to ktCO2
        df_plot = df_plot / 1000.0
        
        # Categorization logic
        categories = {
            'Electricity': ['ELEC', 'ELECTRICITE', 'ELECTRICITÉ'],
            'Hydrogen': ['H2', 'HYDROGEN', 'HYDROGÈNE'],
            'Natural Gas': ['GAS', 'GAZ', 'FUEL'],
            'Other': []
        }
        
        cat_df = pd.DataFrame(index=df_plot.index)
        used_cols = set()
        
        for cat_name, keywords in categories.items():
            if cat_name == 'Other': continue
            matching_cols = []
            for col in df_plot.columns:
                name_upper = str(self.data.resources[col].name if col in self.data.resources else col).upper()
                if any(k in name_upper for k in keywords):
                    matching_cols.append(col)
                    used_cols.add(col)
            if matching_cols:
                cat_df[cat_name] = df_plot[matching_cols].sum(axis=1)
        
        # Add remaining columns to 'Other'
        remaining_cols = [c for c in df_plot.columns if c not in used_cols]
        if remaining_cols:
            cat_df['Other'] = df_plot[remaining_cols].sum(axis=1)

        # Plot stacked area chart with a premium corporate palette
        premium_palette = ['#2C3E50', '#E67E22', '#2980B9', '#8E44AD', '#16A085', '#D35400']
        ax = cat_df.plot.area(stacked=True, alpha=0.85, colormap='tab10' if len(cat_df.columns) > 6 else None)
        if len(cat_df.columns) <= 6:
            for i, poly in enumerate(ax.get_children()):
                if isinstance(poly, plt.Polygon):
                    if i < len(premium_palette):
                        poly.set_facecolor(premium_palette[i])

        plt.title('INDIRECT EMISSIONS BREAKDOWN (SCOPE 2 & 3)', fontsize=15, weight='bold', pad=20)
        plt.ylabel('ktCO2', fontsize=12, weight='semibold')
        plt.xlabel('Year', fontsize=12, weight='semibold')
        
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), title="Category", title_fontsize='12', ncol=3, fontsize=12)
        self._apply_premium_style(ax)

        plt.tight_layout()
        fig.subplots_adjust(bottom=0.2)
        self._add_watermark(fig)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, f'{self.scenario_name}_Indirect_Emissions.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        # Store data
        self.charts_data.append(("Indirect Emissions Breakdown by Category (Scope 2 & 3)", cat_df))
        
    def _plot_investment_costs(self, df: pd.DataFrame, df_finance: pd.DataFrame = None):
        """Plot the Investment Plan, focusing on Implementation Costs (M€) and budget limits."""
        # --- Charting Data (Use Granular Projects) ---
        df_plot_raw = self.df_projects.copy()
        
        # We look for CAPEX columns (raw names from df_projects, not ending in ##tCO2 or _labels or _is_new)
        excluded_suffixes = ('##tCO2', '_labels', '_is_new', 'Financing Interests', 'Year', 'Yearly_Total')
        capex_cols = [c for c in df_plot_raw.columns if not any(c.endswith(s) for s in excluded_suffixes) and c != 'Year']
        
        df_plot = df_plot_raw[['Year'] + capex_cols].set_index('Year')
        # Convert to Millions of Euros (M€)
        df_plot = df_plot / 1_000_000.0
        
        # Rename columns for internal logic (already using raw IDs here)
        
        # Load metadata
        process_labels = {} 
        for col in df_plot_raw.columns:
            if col.endswith('_labels'):
                p_key = col[:-7]
                process_labels[p_key] = df_plot_raw.set_index('Year')[col].to_dict()

        # Filter out purely zero columns (threshold in M€)
        df_bars = df_plot.loc[:, (df_plot.sum(axis=0) > 1e-6)]
        
        years = list(df_plot.index)
        n_years = len(years)
        x_idx = np.arange(n_years)
        
        # Layout
        fig, ax_main = plt.subplots(figsize=(12, 9))
        fig.set_facecolor('#FBFBFB')
        ax_main.set_facecolor('#FBFBFB')
        
        custom_handles = []
        custom_labels = []

        if not df_bars.empty:
            # Corporate palette for processes
            proc_palette = [
                '#004B87', '#007AA5', '#0095A8', '#00A69C', '#24A148',
                '#63BA3C', '#BFD02C', '#FADA00', '#F8B400', '#F08000'
            ]
            
            # Bright / contrasting palette for borders (Technologies)
            tech_palette = [
                '#D0021B', '#F5A623', '#F8E71C', '#8B572A', '#BD10E0',
                '#9013FE', '#4A90E2', '#E24A8D', '#00FF00', '#FF00FF'
            ]
            
            # Map processes to colors
            proc_ids = sorted(list({col.split('##')[0] for col in df_bars.columns if '##' in col}))
            proc_to_color = {pid: proc_palette[i % len(proc_palette)] for i, pid in enumerate(proc_ids)}
            
            # Map technologies to borders
            tech_ids = sorted(list({col.split('##')[1] if '##' in col else col.split('##')[0] for col in df_bars.columns if '##' in col}))
            tech_to_edge_color = {tid: tech_palette[i % len(tech_palette)] for i, tid in enumerate(tech_ids)}

            colors_for_text = []

            # Shadows
            df_plot['Yearly_Total'] = df_bars.sum(axis=1)
            max_h_shadow = df_plot['Yearly_Total'].max()
            for xi, total in enumerate(df_plot['Yearly_Total']):
                if total > 0:
                    shadow = plt.Rectangle((xi - 0.48, -max_h_shadow*0.005), 0.96, total,
                                         color='black', alpha=0.03, zorder=1)
                    ax_main.add_patch(shadow)

            # Build Custom Legend handles
            import matplotlib.patches as mpatches
            
            # Header entry for the legend explanation
            custom_handles.append(mpatches.Patch(color='none'))
            custom_labels.append("Investment Plan (Implementation Costs)")
            
            for pid in proc_ids:
                if pid == 'INDIRECT':
                    proc_name = "Indirect Tech (DAC/Credits)"
                elif hasattr(self.opt, 'entity') and hasattr(self.opt.entity, 'processes') and pid in self.opt.entity.processes:
                    proc_name = getattr(self.opt.entity.processes[pid], 'name', pid)
                else:
                    proc_name = pid
                custom_handles.append(mpatches.Patch(facecolor=proc_to_color[pid], edgecolor='none'))
                custom_labels.append(f"Process: {proc_name}")
                
            for tid in tech_ids:
                tech_name = self.data.technologies[tid].name if tid in self.data.technologies else tid
                custom_handles.append(mpatches.Patch(facecolor='none', edgecolor=tech_to_edge_color[tid], linewidth=3.0))
                custom_labels.append(f"Tech: {tech_name}")

            # Bar plot
            bottoms = np.zeros(n_years)
            for i, col in enumerate(df_bars.columns):
                parts = col.split('##')
                pid = parts[0]
                tid = parts[1] if len(parts) > 1 else pid
                fill_color = proc_to_color[pid]
                edge_color = tech_to_edge_color[tid]
                
                colors_for_text.append(fill_color)
                
                vals = df_bars[col].values
                edge_colors = [edge_color if v > 1e-4 else 'none' for v in vals]
                
                ax_main.bar(x_idx, df_bars[col], bottom=bottoms, color=fill_color,
                               alpha=0.92, width=0.92, edgecolor=edge_colors, linewidth=2.0, zorder=3)
                bottoms += df_bars[col].values

            # Unit Labels
            cumul_bott = np.zeros(n_years)
            for i, col in enumerate(df_bars.columns):
                vals = df_bars[col].values
                for xi, val in enumerate(vals):
                    yr = years[xi]
                    y_bot = cumul_bott[xi]
                    y_mid = y_bot + val / 2.0
                    
                    if val > 1e-4:
                        label = process_labels.get(col, {}).get(yr, "")
                        if label:
                            import matplotlib.colors as mcolors
                            c = colors_for_text[i]
                            lumi = 0.299 * mcolors.to_rgb(c)[0] + 0.587 * mcolors.to_rgb(c)[1] + 0.114 * mcolors.to_rgb(c)[2]
                            text_color = 'white' if lumi < 0.6 else '#333333'
                            
                            bbox_props = None
                            if "DAC" in label.upper() or "CREDIT" in label.upper() or "INDIRECT" in label.upper():
                                bbox_props = dict(facecolor='black', alpha=0.85, edgecolor='none', boxstyle='round,pad=0.2')
                                text_color = 'white'
                            
                            ax_main.text(xi, y_mid, label, ha='center', va='center', 
                                         fontsize=8, color=text_color, weight='bold', zorder=25,
                                         bbox=bbox_props)
                    cumul_bott[xi] += val

            # --- Investment Limits ---
            # Extract limits from df (converted to M€)
            if 'Budget_Limit' in self.df_costs.columns:
                budget_lim = self.df_costs.set_index('Year')['Budget_Limit'] / 1_000_000.0
                total_lim = self.df_costs.set_index('Year')['Total_Limit'] / 1_000_000.0
                
                # Plot Budget Limit (Cash Propre)
                ax_main.step(x_idx, budget_lim, where='mid', color='#2ECC71', linestyle='--', linewidth=2.5, 
                             label='Self-funded Limit (Own Cash)', zorder=10)
                custom_handles.append(plt.Line2D([0], [0], color='#2ECC71', linestyle='--', linewidth=2.5))
                custom_labels.append('Self-funded Limit (Own Cash)')
                
                # Plot Total Limit (Loans included)
                ax_main.step(x_idx, total_lim, where='mid', color='#E74C3C', linestyle='-.', linewidth=2.5, 
                             label='Total Investment Limit (Incl. Loans)', zorder=11)
                custom_handles.append(plt.Line2D([0], [0], color='#E74C3C', linestyle='-.', linewidth=2.5))
                custom_labels.append('Total Investment Limit (Incl. Loans)')

        # ── Main axis labels & grid ──────────────────────────────────────────────
        ax_main.set_title('Investment Plan: Implementation Costs (M€)', 
                         fontsize=16, weight='bold', pad=30, color='#1A1A1A')
        ax_main.set_ylabel('Annual Implementation Cost (M€)', fontsize=12, weight='semibold', color='#333333')
        
        ax_main.grid(axis='y', color='#E0E0E0', linestyle='-', linewidth=0.5, alpha=0.8, zorder=0)
        for spine in ['top', 'right']:
            ax_main.spines[spine].set_visible(False)
        for spine in ['left', 'bottom']:
            ax_main.spines[spine].set_color('#CCCCCC')
            ax_main.spines[spine].set_linewidth(0.8)
        
        # Show years per 5-year steps
        ticks_5yr = [xi for xi, y in enumerate(years) if y % 5 == 0]
        ax_main.set_xticks(ticks_5yr)
        ax_main.set_xticklabels([str(years[xi]) for xi in ticks_5yr], rotation=0, ha='center', fontsize=11, color='#444444', weight='semibold')
        ax_main.set_xlim(-0.5, n_years - 0.5)
        
        # Legend with description
        explanation = "This chart displays the annual implementation costs (CAPEX) for new projects, compared against investment limits."
        
        leg = ax_main.legend(custom_handles, custom_labels,
                       loc='upper center', bbox_to_anchor=(0.5, -0.18), 
                       ncol=min(3, len(custom_labels)), 
                       fontsize=12, frameon=True, title=explanation, title_fontsize=11)
        
        self._apply_premium_style(ax_main)
        
        plt.tight_layout()
        fig.subplots_adjust(bottom=0.28) # Make room for the large legend
        self._add_watermark(fig)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, f'{self.scenario_name}_Investment_Plan_Costs.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        # Store data for Excel
        self.charts_data.append(("Investment Plan: Implementation Costs", df_plot))

    def _plot_external_financing(self, df_costs: pd.DataFrame, df_finance: pd.DataFrame):
        """Standardizes the Public Aids chart to include private bank loans (Financing)."""
        df_costs = df_costs.copy()
        if 'Year' in df_costs.columns:
            df_costs.set_index('Year', inplace=True)
            
        # 1. Public Aids
        aid_cols = [c for c in df_costs.columns if c.startswith('Aid_')]
        df_aids = df_costs[aid_cols].copy()
        df_aids = df_aids.loc[:, (df_aids != 0).any(axis=0)]
        df_aids = df_aids / 1_000_000.0 # Convert to M€
        
        # 2. Private Financing (Bank Loans)
        df_fin = df_finance.set_index('Year')
        private_loans = df_fin['Loan_Principal_Taken (M€)']
        
        # Merge for plotting
        df_plot = df_aids.copy()
        if private_loans.any():
            df_plot['Private Bank Loans'] = private_loans

        if df_plot.empty:
            self._save_no_data_chart(
                f'{self.scenario_name}_Financing.png',
                'FINANCING STRATEGY',
                'No public aids or private loans were triggered in this scenario.'
            )
            return
            
        # Design a dark-themed, highly aesthetic chart
        fig, ax = plt.subplots(figsize=(12, 9), facecolor='#0D0D14')
        ax.set_facecolor('#0D0D14')
        
        years = list(df_plot.index)
        
        # High-contrast, neon-like color palette
        colors = ['#00E5FF', '#FF007F', '#00FF7F', '#FFD700', '#FF8C00', '#9D00FF', '#FF00FF', '#CCFF00']
        
        # Rename columns to show Grant/CCfD clearly for public aids
        rename_map = {}
        for col in df_plot.columns:
            if col.startswith('Aid_'):
                parts = col.split('_', 2)
                if len(parts) == 3:
                    t_name = self.data.technologies[parts[2]].name if parts[2] in self.data.technologies else parts[2]
                    rename_map[col] = f"{parts[1]} - {t_name}"
            else:
                rename_map[col] = col
        df_plot = df_plot.rename(columns=rename_map)
        
        # Plot stacked bars
        pos_bottoms = [0.0] * len(years)
        neg_bottoms = [0.0] * len(years)
        for i, col in enumerate(df_plot.columns):
            c = colors[i % len(colors)]
            values = df_plot[col].values
            
            # Determine starting points for this stack (positive vs negative)
            current_bottoms = []
            for v_idx, v in enumerate(values):
                if v >= 0:
                    current_bottoms.append(pos_bottoms[v_idx])
                    pos_bottoms[v_idx] += v
                else:
                    current_bottoms.append(neg_bottoms[v_idx])
                    neg_bottoms[v_idx] += v
            
            # Sublte glow effect on bars
            ax.bar(years, values, bottom=current_bottoms, color=c, alpha=0.9, 
                   edgecolor=c, linewidth=1.5, label=col, width=0.55, zorder=3)
            # Core bar highlight
            ax.bar(years, values, bottom=current_bottoms, color='white', alpha=0.15, 
                   width=0.55, zorder=3)
                   
        # Cumulative Line (Net Total over time)
        yearly_net = [p + n for p, n in zip(pos_bottoms, neg_bottoms)]
        cumul_sum = []
        c_val = 0
        for v in yearly_net:
            c_val += v
            cumul_sum.append(c_val)
            
        ax2 = ax.twinx()
        glow_color = '#00ffcc'
        for alpha, lw in zip([0.05, 0.1, 0.2, 0.4, 0.7], [18, 12, 8, 4, 2]):
            ax2.plot(years, cumul_sum, color=glow_color, alpha=alpha, linewidth=lw, zorder=4)
            
        ax2.plot(years, cumul_sum, color='#FFFFFF', linewidth=2.0, marker='o', 
                 markersize=7, markerfacecolor=glow_color, markeredgecolor='#FFFFFF', 
                 markeredgewidth=1.5, label='Cumulative Total (M€)', zorder=5)
                 
        # Text annotations on the cumulative line
        for i, (yr, val) in enumerate(zip(years, cumul_sum)):
            if i % 5 == 0 or i == len(years)-1:
                if val > 0.1:
                    ax2.annotate(f"{val:.1f} M€", xy=(yr, val), xytext=(0, 20), 
                                 textcoords="offset points", ha='center', color=glow_color,
                                 fontsize=10, weight='bold', 
                                 bbox=dict(boxstyle="round,pad=0.3", fc="#0D0D14", ec=glow_color, alpha=0.7),
                                 arrowprops=dict(arrowstyle="-", color=glow_color, alpha=0.8), zorder=6)
        
        # Styling details
        fig.suptitle("FINANCING STRATEGY", 
                     color='white', fontsize=18, weight='bold', fontfamily='sans-serif', y=0.96)
        ax.set_title("Annual support from public aids (grants/CCfD) and private loans", 
                     color='#AAAAAA', fontsize=12, pad=10)
                     
        ax.set_ylabel("Annual Triggered Support (M€)", color='white', fontsize=12, labelpad=10)
        ax2.set_ylabel("Cumulative Support (M€)", color=glow_color, fontsize=12, labelpad=10, weight='bold')
        
        self._apply_premium_style(ax, is_dark=True)
        
        for spine in ax2.spines.values():
            spine.set_visible(False)
            
        # Custom Legends at the BOTTOM
        lines_1, labels_1 = ax.get_legend_handles_labels()
        lines_2, labels_2 = ax2.get_legend_handles_labels()
        if lines_1 or lines_2:
            leg = ax.legend(lines_1 + lines_2, labels_1 + labels_2,
                      loc='upper center', bbox_to_anchor=(0.5, -0.15), 
                      facecolor='#0D0D14', edgecolor='#2B2B36', labelcolor='white', fontsize=14, 
                      framealpha=0.9, borderpad=1, ncol=min(3, len(lines_1)+len(lines_2)))
            for text in leg.get_texts():
                text.set_weight("bold")
                
        plt.tight_layout()
        fig.subplots_adjust(bottom=0.22)
        self._add_watermark(fig, is_dark_bg=True)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        try:
            # We construct the path absolutely and explicitly to avoid invisible unicode characters
            f_name = f'{self.scenario_name}_Financing.png'
            out_file = os.path.join(self.results_dir, f_name)
            plt.savefig(out_file, dpi=300, bbox_inches='tight', facecolor='#0D0D14')
        except Exception as e:
            print(f"  [red][Reporter][/red] [!] Error saving {f_name}: {e}")
        plt.close()

        # Store data
        self.charts_data.append(("Financing Strategy: Public & Private", df_plot))

    def _plot_total_opex(self, df_cons: pd.DataFrame, df_emis: pd.DataFrame):
        df_plot = df_cons.copy()
        df_plot.set_index('Year', inplace=True)
        years = list(df_plot.index)
        
        # 1. Resource Costs (grouped by resource ID)
        res_costs = {}
        for res_id in df_plot.columns:
            if res_id in ['EN_H2_ON_SITE']: continue # Avoid double counting
            res_costs[res_id] = []
            for t in years:
                cons = df_plot.at[t, res_id]
                price = self.data.time_series.resource_prices.get(res_id, {}).get(t, 0.0)
                if price == 0.0 and ('H2' in res_id.upper() or 'HYDROGEN' in res_id.upper()):
                    price = self.data.time_series.resource_prices.get('EN_GREY_H2_C', {}).get(t, 0.0)
                cost = (cons * price if cons > 0 else 0) / 1_000_000.0
                res_costs[res_id].append(cost)
        
        df_res_costs = pd.DataFrame(res_costs, index=years)
        # Rename resource columns for humans
        res_rename = {}
        for col in df_res_costs.columns:
            if col in self.data.resources:
                res_name = self.data.resources[col].name
                res_rename[col] = res_name.upper() if res_name else col.upper()
        df_res_costs.rename(columns=res_rename, inplace=True)
        # Filter zero resources
        df_res_costs = df_res_costs.loc[:, (df_res_costs.abs() > 1e-4).any(axis=0)]
        
        # 2. Technology Individual OPEX
        tech_opex_details = {}
        # We look for all technologies that might be implemented
        for t_id, tech in self.data.technologies.items():
            if t_id == 'UP': continue
            tech_annual_opex = []
            is_active = False
            for t in years:
                yr_val = 0.0
                for p_id, process in self.opt.entity.processes.items():
                    if t_id in process.valid_technologies:
                        act_v = getattr(self.opt.active_vars[(t, p_id, t_id)], 'varValue', 0.0) or 0.0
                        if act_v > 0.01:
                            is_active = True
                            cap_opex = 1.0
                            if tech.opex_per_unit:
                                if tech.opex_unit == 'tCO2': cap_opex = self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                                elif 'MW' in str(tech.opex_unit).upper():
                                    cap_opex = (self.opt.entity.base_consumptions.get('EN_FUEL', 0.0) * process.consumption_shares.get('EN_FUEL', 0.0)) / self.opt.entity.annual_operating_hours
                            current_opex = tech.opex_by_year.get(t, tech.opex)
                            yr_val += (current_opex * cap_opex / process.nb_units) * act_v
                tech_annual_opex.append(yr_val / 1_000_000.0)
            
            if is_active:
                col_name = f"{tech.name.upper()} OPEX" if tech.name else f"{t_id.upper()} OPEX"
                tech_opex_details[col_name] = tech_annual_opex

        df_tech_costs = pd.DataFrame(tech_opex_details, index=years)
        
        # 3. DAC OPEX
        dac_opex = []
        if self.data.dac_params.active:
            for t in years:
                dac_total_v = getattr(self.opt.dac_total_capacity_vars.get(t), 'varValue', 0.0) or 0.0
                dac_opex.append((dac_total_v * self.data.dac_params.opex_by_year.get(t, 0.0)) / 1_000_000.0)
        
        # 4. Carbon Tax (Gross)
        tax_costs = df_emis.set_index('Year')['Tax_Cost_MEuros'].tolist()
        
        # 5. Carbon Credits
        credit_costs = []
        for t in years:
            cred_v = getattr(self.opt.credit_purchased_vars.get(t), 'varValue', 0.0) or 0.0
            credit_costs.append((cred_v * self.data.credit_params.cost_by_year.get(t, 0.0)) / 1_000_000.0)
            
        # 6. CCS Storage & Transport (Specific calculated OPEX)
        ccs_st_costs = []
        for t in years:
            yr_ccs_st = 0.0
            for p_id, process in self.opt.entity.processes.items():
                for t_id in process.valid_technologies:
                    if "CCS" in t_id.upper() or "CCU" in t_id.upper():
                        tech = self.data.technologies[t_id]
                        imp = tech.impacts.get('CO2_EM')
                        if imp and (imp['type'] == 'variation' or imp['type'] == 'up'):
                            act_v = getattr(self.opt.active_vars[(t, p_id, t_id)], 'varValue', 0.0) or 0.0
                            if act_v > 0.01:
                                reduction_frac_per_unit = (-imp['value'] if imp['value'] < 0 else 0) / process.nb_units
                                max_emis_for_proc = self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                                captured_tons_per_unit = reduction_frac_per_unit * max_emis_for_proc
                                
                                s_price = self.data.time_series.resource_prices.get('CO2_STORAGE', {}).get(t, 0.0)
                                tr_price = self.data.time_series.resource_prices.get('CO2_TRANSPORT', {}).get(t, 0.0)
                                yr_ccs_st += (s_price + tr_price) * captured_tons_per_unit * act_v
            ccs_st_costs.append(yr_ccs_st / 1_000_000.0)

        # Combine into a plotting dataframe in a clean order
        df_bars = pd.concat([df_res_costs, df_tech_costs], axis=1)
        
        if sum(ccs_st_costs) > 1e-4:
            df_bars['CCS STORAGE & TRANSPORT'] = ccs_st_costs
            
        if sum(dac_opex) > 1e-4:
            df_bars['DAC OPEX'] = dac_opex
        
        df_bars['CARBON TAX (GROSS)'] = tax_costs
        
        if sum(credit_costs) > 1e-4:
            df_bars['CARBON CREDITS'] = credit_costs
            
        # Final clean-up of empty columns
        df_bars = df_bars.loc[:, (df_bars.abs() > 1e-4).any(axis=0)]
        total_opex = df_bars.sum(axis=1)
        
        # Add Total to the dataframe for Excel export (before plotting)
        df_bars_export = df_bars.copy()
        df_bars_export['TOTAL ANNUAL OPEX (M€)'] = total_opex
        
        fig, ax = plt.subplots(figsize=(14, 10))
        # Use tab20 colors for variety
        df_bars.plot.area(ax=ax, stacked=True, alpha=0.85, colormap='tab20')
        
        # Plot Total Line
        ax.plot(years, total_opex, color='#2C3E50', linestyle='--', linewidth=3, 
                marker='o', markersize=7, markerfacecolor='white', label='Total Annual OPEX (M€)')
        
        # Annotations
        for t in [years[0], years[-1]]:
            val = total_opex.loc[t]
            ax.annotate(f"{val:.1f} M€", xy=(t, val), xytext=(0, 15), 
                        textcoords="offset points", ha='center', va='bottom', fontsize=11,
                        weight='bold', color='#2C3E50',
                        bbox=dict(boxstyle='round,pad=0.3', fc='white', ec='#2C3E50', alpha=0.9))
        
        ax.set_title("TOTAL OPEX: ANNUAL OPERATIONAL EXPENDITURE BREAKDOWN", 
                     fontsize=18, weight='bold', pad=30)
        ax.set_ylabel("Annual OPEX (M€)", fontsize=13, weight='semibold')
        ax.set_xlabel("Year", fontsize=13, weight='semibold')
        ax.set_xticks(years)
        ax.set_xticklabels([str(y) for y in years], rotation=0)
        
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), fontsize=10, ncol=3, frameon=True)
        self._apply_premium_style(ax)
        
        plt.tight_layout()
        self._add_watermark(fig)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, f'{self.scenario_name}_Total_OPEX.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        self.charts_data.append(("Generalized Operational Expenditure Breakdown", df_bars_export))

    def _plot_carbon_tax_and_avoided(self, df: pd.DataFrame):
        df_plot = df.copy()
        years = df_plot['Year'].tolist()
        
        # Calculate Avoided Costs in MEuros
        df_plot['Avoided_Cost_Reduced_MEuros'] = df_plot['Really_Avoided_CO2_kt'] * 1000 * df_plot['Tax_Price'] / 1_000_000.0
        df_plot['Avoided_Cost_Total_MEuros'] = df_plot['Avoided_Direct_CO2_kt'] * 1000 * df_plot['Tax_Price'] / 1_000_000.0
        
        # Invert avoided costs for downward display on same axis
        df_plot['Avoided_Cost_Reduced_MEuros_Neg'] = -df_plot['Avoided_Cost_Reduced_MEuros']
        df_plot['Avoided_Cost_Total_MEuros_Neg'] = -df_plot['Avoided_Cost_Total_MEuros']

        # Premium Financial Balance Visualization
        fig, ax = plt.subplots(figsize=(12, 9))
        
        # ── POSITIVE REGION : Costs ──
        # Vertical dotted lines with crosses for annual taxes
        markerline_tax, stemlines_tax, baseline_tax = ax.stem(years, df_plot['Tax_Cost_MEuros'], linefmt='#d62728', markerfmt='x', basefmt=' ')
        plt.setp(stemlines_tax, linestyle=':', linewidth=1.5, color='#d62728')
        plt.setp(markerline_tax, markersize=8, color='#d62728', markeredgewidth=1.5)
        markerline_tax.set_label('Gross Carbon Tax (M€)')

        # CCfD Refund (State compensation)
        if 'CCfD_Refund_MEuros' in df_plot.columns and df_plot['CCfD_Refund_MEuros'].abs().max() > 1e-4:
            ax.scatter(years, df_plot['CCfD_Refund_MEuros'], color='tab:orange', marker='D', s=40, edgecolors='white', linewidths=0.5, label='CCfD State Refund (M€)', zorder=6)
            
        # Net Carbon Cost (Tax - Refund)
        if 'Net_Tax_Cost_MEuros' in df_plot.columns:
            ax.plot(years, df_plot['Net_Tax_Cost_MEuros'], color='#d62728', linewidth=2.5, linestyle=':', label='Net Annual Cost (M€)', zorder=7)

        # ── NEGATIVE REGION : Avoided Costs ──
        # Plot the two levels as circles/dotted pointing DOWN
        ax.plot(years, df_plot['Avoided_Cost_Reduced_MEuros_Neg'], color='#2ca02c', linestyle=':', marker='o', markersize=4, label='Avoided (Reduced at Source) (M€)', zorder=5)
        ax.plot(years, df_plot['Avoided_Cost_Total_MEuros_Neg'], color='#17becf', linestyle=':', marker='o', markersize=4, label='Total Avoided & Captured (M€)', zorder=5)
        
        # Add orthogonal segments
        ax.vlines(years, 0, df_plot['Avoided_Cost_Reduced_MEuros_Neg'], color='#2ca02c', linestyle='-', linewidth=1.5, alpha=0.4, zorder=4)
        ax.vlines(years, df_plot['Avoided_Cost_Reduced_MEuros_Neg'], df_plot['Avoided_Cost_Total_MEuros_Neg'], color='#17becf', linestyle='-', linewidth=1.5, alpha=0.4, zorder=4)
        
        # Horizontal separation line at 0
        ax.axhline(0, color='black', linewidth=1.2, alpha=0.9, zorder=3)

        # Left Axis Styling
        ax.set_ylabel("Annual Financial Impact (M€)", fontsize=12, weight='bold')
        ax.grid(axis='y', linestyle='--', alpha=0.5)
        ax.grid(axis='x', linestyle='--', alpha=0.3)
        ax.set_title("Carbon Tax and Avoided Costs Balance", fontsize=15, weight='bold')
        ax.tick_params(axis='y', labelsize=10)
        
        # ── RIGHT AXIS : Avoided Costs ──
        # We create a twin axis for symmetry and labels
        ax_right = ax.twinx()
        ax_right.set_ylabel("Avoided Costs (M€)", fontsize=12, weight='bold', color='#2ca02c')
        ax_right.tick_params(axis='y', labelcolor='#2ca02c', labelsize=10)
        
        # Symmetrical limits to center the 0 line
        y_max = df_plot['Tax_Cost_MEuros'].max()
        y_min_abs = df_plot['Avoided_Cost_Total_MEuros'].max()
        ylim_max = max(y_max, y_min_abs) * 1.25
        if ylim_max == 0: ylim_max = 1
        ax.set_ylim(-ylim_max, ylim_max)
        ax_right.set_ylim(-ylim_max, ylim_max)
        
        # Right ticks showing absolute values for clarity
        ax_right.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: f"{abs(x):.0f}"))

        # ── Annotate tax prices ──
        if 'Tax_Price' in df_plot.columns:
            for i, year in enumerate(years):
                if i % 5 == 0 or year == years[-1]:
                    tax_price = df_plot['Tax_Price'].iloc[i]
                    val = df_plot['Tax_Cost_MEuros'].iloc[i]
                    ax.annotate(f"{tax_price:,.0f} €/t", xy=(year, val), xytext=(0, 12), 
                                    textcoords="offset points", ha='center', va='bottom', fontsize=9, weight='bold', color='#d62728',
                                    arrowprops=dict(arrowstyle='-', color='#d62728', lw=0.8, alpha=0.6))
        
        ax.set_title("CARBON TAX AND AVOIDED COSTS BALANCE", fontsize=16, weight='bold', pad=25)
        ax.set_ylabel("Annual Financial Impact (M€)", fontsize=12, weight='semibold')
        ax.set_xlabel("Year", fontsize=12, weight='semibold')
        ax.set_xticks(years)
        ax.set_xticklabels([str(y) for y in years], rotation=45, ha='right')
        
        self._apply_premium_style(ax)
        
        for spine in ax_right.spines.values():
            spine.set_visible(False)
            
        # Symmetrical limits to center the 0 line
        y_max = df_plot['Tax_Cost_MEuros'].max()
        y_min_abs = df_plot['Avoided_Cost_Total_MEuros'].max()
        ylim_max = max(y_max, y_min_abs) * 1.25
        if ylim_max == 0: ylim_max = 1
        ax.set_ylim(-ylim_max, ylim_max)
        ax_right.set_ylim(-ylim_max, ylim_max)
        
        # Combined Legend
        lines_1, labels_1 = ax.get_legend_handles_labels()
        ax.legend(lines_1, labels_1, loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, fontsize=12, frameon=True)
        
        plt.tight_layout()
        self._add_watermark(fig)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, f'{self.scenario_name}_Carbon_Tax_And_Avoided.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        # Store data
        self.charts_data.append(("Carbon Tax and Avoided Costs Balance", df_plot))

    def _plot_carbon_prices(self):
        """Generates a stunning visualization of carbon price trajectory and penalty factors."""
        years = self.years
        carbon_prices = [self.data.time_series.carbon_prices.get(y, 0.0) for y in years]
        penalties = [self.data.time_series.carbon_penalties.get(y, 0.0) for y in years]
        effective_prices = [p * (1.0 + x) for p, x in zip(carbon_prices, penalties)]

        # Premium Dark Theme
        fig, ax = plt.subplots(figsize=(12, 9), facecolor='#0D0D14')
        ax.set_facecolor('#0D0D14')

        # Glowy colors
        base_color = '#00CCFF'  # Cyan
        penalty_color = '#FF007F' # Neon Pink
        effective_color = '#CCFF00' # Lime

        # Plot Base Carbon Price with glow
        for alpha, lw in zip([0.05, 0.1, 0.2, 0.4], [15, 10, 6, 2]):
            ax.plot(years, carbon_prices, color=base_color, alpha=alpha, linewidth=lw, zorder=3)
        ax.plot(years, carbon_prices, color=base_color, linewidth=2, label='Market Carbon Price', zorder=4)

        # If penalties exist, plot them too
        if any(x > 0 for x in penalties):
            # Plot Effective Price (with penalty)
            for alpha, lw in zip([0.05, 0.1, 0.2, 0.4], [15, 10, 6, 2]):
                ax.plot(years, effective_prices, color=effective_color, alpha=alpha, linewidth=lw, zorder=5)
            ax.plot(years, effective_prices, color=effective_color, linestyle='--', linewidth=2, 
                    label='Penalty', zorder=6)
            
            # Fill the penalty gap
            ax.fill_between(years, carbon_prices, effective_prices, color=penalty_color, alpha=0.15, 
                            label='Penalty Gap', hatch='//', zorder=2)

        # --- NEW: Plot Carbon Strike Prices for active CCfD contracts ---
        strike_color = '#FFD700' # Gold/Yellow neon
        if hasattr(self.opt, 'ccfd_used_vars'):
            strike_plotted = False
            for (t_inv, p_id, t_id), var in self.opt.ccfd_used_vars.items():
                if var.varValue and var.varValue > 0.5:
                    tech = self.data.technologies[t_id]
                    ccfd_p = self.data.ccfd_params
                    
                    # Calculate strike price at year of investment
                    base_p = self.data.time_series.carbon_prices.get(t_inv, 0.0)
                    strike_val = (1.0 + ccfd_p.eua_price_pct) * base_p
                    
                    # Contract period
                    start_yr = t_inv + tech.implementation_time
                    end_yr = start_yr + ccfd_p.duration
                    
                    # Plot horizontal line across the contract period
                    contract_years = [y for y in years if start_yr <= y < end_yr]
                    if contract_years:
                        label = 'Contractual Strike' if not strike_plotted else None
                        ax.hlines(y=strike_val, xmin=min(contract_years), xmax=max(contract_years), 
                                  color=strike_color, linestyle='-.', linewidth=2.5, 
                                  label=label, zorder=7)
                        
                        # Add small label above the line
                        ax.text(min(contract_years), strike_val + 5, f"Strike: {tech.name}", 
                                color=strike_color, fontsize=8, weight='bold', alpha=0.9)
                        strike_plotted = True

        # Styling
        ax.set_title("CARBON QUOTA PRICE TRAJECTORY", color='white', fontsize=18, weight='bold', pad=30)
        ax.set_ylabel("Euro (€) / tCO2", color='white', fontsize=12, labelpad=15)
        ax.set_xlabel("Year", color='white', fontsize=12, labelpad=15)
        
        self._apply_premium_style(ax, is_dark=True)

        # Annotations for start/end and every 5 years
        for i, y in enumerate(years):
            if i == 0 or i == len(years) - 1 or y % 5 == 0:
                p = carbon_prices[i]
                ep = effective_prices[i]
                
                # Base price annotation
                ax.annotate(f"{p:.1f} €", xy=(y, p), xytext=(0, 12), textcoords="offset points",
                             color=base_color, weight='bold', ha='center', fontsize=9,
                             bbox=dict(boxstyle='round,pad=0.2', fc='#0D0D14', ec=base_color, alpha=0.7, lw=0.5))
                
                # Effective price annotation (only if different)
                if ep > p + 0.01:
                    ax.annotate(f"{ep:.1f} €", xy=(y, ep), xytext=(0, 20), textcoords="offset points",
                                 color=effective_color, weight='bold', ha='center', fontsize=9,
                                 bbox=dict(boxstyle='round,pad=0.2', fc='#0D0D14', ec=effective_color, alpha=0.7, lw=0.5))

        # Legend
        leg = ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.18), facecolor='#0D0D14', 
                        edgecolor='#2B2B36', labelcolor='white', fontsize=14, ncol=3)
        if leg:
            for text in leg.get_texts():
                text.set_weight("bold")

        plt.tight_layout()
        fig.subplots_adjust(bottom=0.22)
        self._add_watermark(fig, is_dark_bg=True)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, f'{self.scenario_name}_Carbon_Prices_Detailed.png'), dpi=300, 
                    bbox_inches='tight', facecolor='#0D0D14')
        plt.close()
        
        # Store data
        df_cp = pd.DataFrame({
            'Year': years,
            'Market_Price': carbon_prices,
            'Effective_Price': effective_prices,
            'Penalty_Factor': penalties
        }).set_index('Year')
        self.charts_data.append(("Carbon Price & Policy Trajectory", df_cp))


    def _plot_transition_costs(self, df_costs: pd.DataFrame, df_finance: pd.DataFrame, df_emis: pd.DataFrame):
        """Plots the stacked cumulative costs and savings of the ecological transition."""
        df_costs = df_costs.copy()
        if 'Year' in df_costs.columns:
            df_costs.set_index('Year', inplace=True)
            
        # 1. Baseline Calculation ("Nothing Done")
        baseline_data = []
        for t in self.years:
            yr_idx = list(self.years).index(t)
            b_emissions = 0.0
            b_consumptions = {res_id: 0.0 for res_id in self.data.resources}
            
            for p_id, process in self.opt.entity.processes.items():
                up_rate = 0.0
                if 'UP' in process.valid_technologies:
                    up_tech = self.data.technologies['UP']
                    up_imp = up_tech.impacts.get('ALL') or up_tech.impacts.get('CO2_EM')
                    if up_imp and (up_imp['type'] == 'variation' or up_imp['type'] == 'up'):
                        up_rate = abs(up_imp['value'])
                
                up_factor = (1.0 - up_rate) ** yr_idx
                p_base_emis = (self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)) * up_factor
                b_emissions += p_base_emis
                for res_id in self.data.resources:
                    if res_id != 'CO2_EM':
                        p_base_cons = (self.opt.entity.base_consumptions.get(res_id, 0.0) * process.consumption_shares.get(res_id, 0.0)) * up_factor
                        b_consumptions[res_id] += p_base_cons
            
            # Add unallocated portions (fixed)
            allocated_emis_share = sum(p.emission_shares.get('CO2_EM', 0.0) for p in self.opt.entity.processes.values())
            b_emissions += self.opt.entity.base_emissions * (1.0 - allocated_emis_share)
            for res_id in self.data.resources:
                if res_id != 'CO2_EM':
                    allocated_share = sum(p.consumption_shares.get(res_id, 0.0) for p in self.opt.entity.processes.values())
                    b_consumptions[res_id] += self.opt.entity.base_consumptions.get(res_id, 0.0) * (1.0 - allocated_share)
            
            # Carbon Tax Baseline
            tax_price = self.data.time_series.carbon_prices.get(t, 0.0)
            if self.opt.entity.sv_act_mode == "PI":
                fq_pct = self.data.time_series.carbon_quotas_pi.get(t, 0.0)
            else:
                fq_pct = self.data.time_series.carbon_quotas_norm.get(t, 0.0)
            
            taxed_co2_b = b_emissions * (1.0 - fq_pct) if fq_pct <= 1.0 else max(0.0, b_emissions - fq_pct)
            b_tax_cost = taxed_co2_b * tax_price / 1_000_000.0 # M€
            
            # Resource Cost Baseline
            b_res_cost = 0.0
            for res_id, cons_val in b_consumptions.items():
                price = self.data.time_series.resource_prices.get(res_id, {}).get(t, 0.0)
                if price > 0: b_res_cost += cons_val * price
            b_res_cost /= 1_000_000.0 # M€
            
            baseline_data.append({
                'Year': t,
                'Baseline_Tax': b_tax_cost,
                'Baseline_Resource_Cost': b_res_cost
            })
        
        df_b = pd.DataFrame(baseline_data).set_index('Year')
        
        # 2. Delta Dataset Construction (Annual first)
        years = list(df_b.index)
        df_annual = pd.DataFrame(index=years)
        
        # --- POSITIVE COSTS ---
        df_fin = df_finance.set_index('Year')
        # User requested Bank Loan Service to be INTERESTS ONLY
        # Principal repayments are grouped with Self-funded CAPEX to reflect investment effort
        df_annual['Self-funded CAPEX'] = df_fin['Out_of_Pocket_CAPEX (M€)'] + df_fin['Principal_Repayment (M€)']
        df_annual['Bank Loan Service'] = df_fin['Interest_Paid (M€)']
        
        tech_opex = []
        for t in years:
            year_opex = 0.0
            for p_id, process in self.opt.entity.processes.items():
                for t_id in process.valid_technologies:
                    if t_id == 'UP': continue
                    act_v = getattr(self.opt.active_vars[(t, p_id, t_id)], 'varValue', 0.0) or 0.0
                    if act_v > 0.01:
                        tech = self.data.technologies[t_id]
                        cap_opex = 1.0
                        if tech.opex_per_unit:
                            if tech.opex_unit == 'tCO2': cap_opex = self.opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                            elif 'MW' in str(tech.opex_unit).upper():
                                cap_opex = (self.opt.entity.base_consumptions.get('EN_FUEL', 0.0) * process.consumption_shares.get('EN_FUEL', 0.0)) / self.opt.entity.annual_operating_hours
                        current_opex = tech.opex_by_year.get(t, tech.opex)
                        year_opex += (current_opex * cap_opex / process.nb_units) * act_v
            
            if self.data.dac_params.active:
                dac_total_v = getattr(self.opt.dac_total_capacity_vars.get(t), 'varValue', 0.0) or 0.0
                year_opex += dac_total_v * self.data.dac_params.opex_by_year.get(t, 0.0)
            tech_opex.append(year_opex / 1_000_000.0)
        df_annual['Tech & DAC OPEX'] = tech_opex
        
        cred_costs = []
        for t in years:
            cred_v = getattr(self.opt.credit_purchased_vars.get(t), 'varValue', 0.0) or 0.0
            cred_costs.append((cred_v * self.data.credit_params.cost_by_year.get(t, 0.0)) / 1_000_000.0)
        df_annual['Voluntary Carbon Credits'] = cred_costs
        
        # --- NEGATIVE COSTS ---
        public_aids = []
        for t in years:
            grant_total = sum(self.opt.grant_amt_vars.get((t, p_id, t_id)).varValue or 0.0 
                             for p_id, proc in self.opt.entity.processes.items() 
                             for t_id in proc.valid_technologies 
                             if (t, p_id, t_id) in self.opt.grant_amt_vars and self.opt.grant_amt_vars.get((t, p_id, t_id)) is not None and self.opt.grant_amt_vars.get((t, p_id, t_id)).varValue is not None)
            ccfd_refund = df_emis.set_index('Year').at[t, 'CCfD_Refund_MEuros']
            public_aids.append(-(grant_total / 1_000_000.0 + ccfd_refund))
        df_annual['Public Aids (Grants & CCfD)'] = public_aids
        
        actual_tax = df_emis.set_index('Year')['Tax_Cost_MEuros']
        df_annual['Avoided Carbon Tax'] = actual_tax - df_b['Baseline_Tax']
        
        actual_res_cost = []
        for t in years:
            r_cost = 0.0
            for res_id in self.data.resources:
                if res_id == 'CO2_EM': continue
                cons_val = self.opt.cons_vars[(t, res_id)].varValue or 0.0
                price = self.data.time_series.resource_prices.get(res_id, {}).get(t, 0.0)
                if price > 0: r_cost += cons_val * price
            actual_res_cost.append(r_cost / 1_000_000.0)
        df_annual['Resource Savings'] = np.array(actual_res_cost) - df_b['Baseline_Resource_Cost']
        
        # 3. Hybrid Transformation
        df_plot = df_annual.copy() # Areas use annual values
        
        # Rename Resource Savings
        if 'Resource Savings' in df_plot.columns:
            df_plot = df_plot.rename(columns={'Resource Savings': 'Resource Mix Change'})
        
        # Cumulative Balance: sum(Costs - abs(Benefits)).cumsum()
        # Since Benefits are negative in df_annual, sum(axis=1) is (Costs + Benefits) = (Costs - |Benefits|)
        # Calculate on df_annual to ensure correct signs
        df_net_cumul = df_annual.sum(axis=1).cumsum()

        # ── TRANSITION EFFORTS (Positive) ──
        pos_cols = ['Self-funded CAPEX', 'Bank Loan Service', 'Tech & DAC OPEX', 'Voluntary Carbon Credits']
        
        # ── TRANSITION BENEFITS & SAVINGS (Negative) ──
        neg_cols = ['Public Aids (Grants & CCfD)', 'Avoided Carbon Tax']
        
        # Determine Resource Mix Change coloring (Blue if avg > 0 (cost), Green if avg <= 0 (saving))
        res_mean = df_plot['Resource Mix Change'].mean() if 'Resource Mix Change' in df_plot.columns else 0
        if res_mean > 1e-3:
            pos_cols.append('Resource Mix Change')
        else:
            neg_cols.append('Resource Mix Change')
            
        # Clean columns to remove near-zero ones
        pos_cols = [c for c in pos_cols if c in df_plot.columns and df_plot[c].abs().sum() > 1e-3]
        neg_cols = [c for c in neg_cols if c in df_plot.columns and df_plot[c].abs().sum() > 1e-3]
        
        # Refined Palette
        # Efforts: Dark Blues / Purples
        colors_efforts = ['#1F3A93', '#2C3E50', '#6741D9', '#9C36B5', '#3E4444'] 
        # Savings/Aids: Fresh Greens / Teals
        colors_savings = ['#16A085', '#27AE60', '#A2D149', '#F1C40F']
        
        x = years
        fig, ax1 = plt.subplots(figsize=(12, 9), facecolor='white')
        
        # Plot stacked areas (Annual) on ax1
        if pos_cols:
            y_pos = df_plot[pos_cols].values.T
            ax1.stackplot(x, y_pos, labels=pos_cols, colors=colors_efforts[:len(pos_cols)], alpha=0.85, zorder=3)
            
        if neg_cols:
            y_neg = df_plot[neg_cols].values.T
            ax1.stackplot(x, y_neg, labels=neg_cols, colors=colors_savings[:len(neg_cols)], alpha=0.85, zorder=3)
            
        # Create secondary axis for Cumulative Line
        ax2 = ax1.twinx()
        
        # Plot Net Cumulative Cost Line on ax2
        lns = ax2.plot(x, df_net_cumul, color='#E74C3C', linewidth=4, 
                 label='Net Transition Balance (Cumulative)', marker='o', markersize=6, zorder=10)
        
        # Aesthetics for Primary Axis
        ax1.axhline(0, color='#333333', linewidth=1.5, zorder=5)
        
        # ── Dynamic Scenario Investment Cap (Annual Reference) ──────────────────
        if self.data.reporting_toggles.investment_cap > 0:
            cap_val = self.data.reporting_toggles.investment_cap
            ax1.axhline(cap_val, color='#E74C3C', linestyle='--', linewidth=1.5, 
                        label=f'Annual Effort Cap ({cap_val} M€)', zorder=4)

        ax1.set_title("ECOLOGICAL TRANSITION: ANNUAL INVESTMENT EFFORTS & SAVINGS", fontsize=18, weight='bold', pad=35)
        ax1.set_ylabel("Annual Variation vs Baseline (M€)", fontsize=12, weight='semibold', color='#2C3E50')
        ax1.set_xlabel("Year", fontsize=12, weight='semibold')
        
        ax1.yaxis.set_major_formatter(plt.FormatStrFormatter('%g M€'))
        ax1.set_xticks(years)
        ax1.set_xticklabels([str(y) for y in years], rotation=0)
        
        # Adjust Y-limit if investment cap is set
        if self.data.reporting_toggles.investment_cap > 0:
            current_max = ax1.get_ylim()[1]
            cap_val = self.data.reporting_toggles.investment_cap
            ax1.set_ylim(top=max(current_max, cap_val * 1.15))
        
        # Aesthetics for Secondary Axis
        ax2.set_ylabel("Net Cumulative Transition Effort (M€)", fontsize=12, weight='semibold', color='#E74C3C')
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:,.0f} M€'))
        ax2.spines['right'].set_color('#E74C3C')
        ax2.tick_params(axis='y', labelcolor='#E74C3C')
        
        # Add labels for start and end of cumulative line on ax2
        for t in [years[0], years[-1]]:
            val = df_net_cumul.at[t]
            ax2.annotate(f"{val:.1f} M€", xy=(t, val), xytext=(0, 15 if val >=0 else -25),
                        textcoords="offset points", ha='center', fontsize=11, weight='bold',
                        color='#E74C3C', bbox=dict(boxstyle='round,pad=0.3', fc='white', ec='#E74C3C', alpha=0.9))

        # Combined Legend
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper center', 
                  bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=True, fontsize=11)
        
        self._apply_premium_style(ax1)
        # ax2 doesn't need premium style as it's paired with ax1
        
        plt.tight_layout()
        fig.subplots_adjust(bottom=0.2)
        self._add_watermark(fig)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, f'{self.scenario_name}_Transition_Costs.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        # Store data
        df_store = df_plot.copy()
        df_store['Net_Cumulative_Cost'] = df_net_cumul
        self.charts_data.append(("Ecological Transition: Annual & Cumulative Balance", df_store))

    def _plot_interest_paid(self, df_finance: pd.DataFrame):
        """Plots the annual and cumulative interest paid for bank loans."""
        df_plot = df_finance.copy()
        if 'Year' in df_plot.columns:
            df_plot.set_index('Year', inplace=True)
            
        if 'Interest_Paid (M€)' not in df_plot.columns or df_plot['Interest_Paid (M€)'].sum() < 1e-4:
            self._save_no_data_chart(
                f'{self.scenario_name}_Interest_Paid.png',
                'BANK LOANS: INTEREST & CONTRACTED AMOUNTS',
                'No loan interest was paid in this scenario.'
            )
            return
            
        fig, ax = plt.subplots(figsize=(12, 9), facecolor='white')
        
        years = list(df_plot.index)
        interest = df_plot['Interest_Paid (M€)'].values
        loans_taken = df_plot['Loan_Principal_Taken (M€)'].values if 'Loan_Principal_Taken (M€)' in df_plot.columns else np.zeros(len(years))
        
        # Plot bar chart for annual interest
        ax.bar(years, interest, color='#E74C3C', alpha=0.85, edgecolor='#C0392B', width=0.6, label='Annual Interest Paid (M€)')
        
        # Plot cumulative line
        cumul_interest = df_plot['Interest_Paid (M€)'].cumsum()
        ax2 = ax.twinx()
        ax2.plot(years, cumul_interest, color='#2C3E50', linewidth=2.5, marker='o', 
                 markersize=6, markerfacecolor='white', label='Cumulative Interest Paid (M€)', zorder=5)

        # Plot Loans Taken on a new left axis (Only points, no line)
        ax3 = ax.twinx()
        ax3.spines['left'].set_position(('outward', 70))
        ax3.spines['left'].set_visible(True)
        ax3.yaxis.set_label_position('left')
        ax3.yaxis.set_ticks_position('left')
        ax3.plot(years, loans_taken, color='#27AE60', marker='s', 
                 markersize=9, markerfacecolor='#27AE60', markeredgecolor='white', 
                 label='Loan Amount Contracted (M€)', linestyle='None')
                 
        # Align all zeros on the bottom
        ax.set_ylim(bottom=0)
        ax2.set_ylim(bottom=0)
        ax3.set_ylim(bottom=0)
                 
        ax.set_title("BANK LOANS: INTEREST & CONTRACTED AMOUNTS", fontsize=16, weight='bold', pad=20)
        ax.set_ylabel("Annual Interest Paid (M€)", fontsize=12, weight='semibold', color='#E74C3C')
        ax2.set_ylabel("Cumulative Interest Paid (M€)", fontsize=12, weight='semibold', color='#2C3E50')
        ax3.set_ylabel("Loan Amount Contracted (M€)", fontsize=12, weight='semibold', color='#27AE60')
        ax.set_xlabel("Year", fontsize=12, weight='semibold')
        
        ax.set_xticks(years)
        ax.set_xticklabels([str(y) for y in years], rotation=45, ha='right')
        
        # Add labels for start and end of cumulative line on ax2
        for t in [years[0], years[-1]]:
            val = cumul_interest.at[t]
            ax2.annotate(f"{val:.1f} M€", xy=(t, val), xytext=(0, 15),
                        textcoords="offset points", ha='center', fontsize=10, weight='bold',
                        color='#2C3E50', bbox=dict(boxstyle='round,pad=0.3', fc='white', ec='#2C3E50', alpha=0.8))

        # Combine legends
        lines1, labels1 = ax.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        lines3, labels3 = ax3.get_legend_handles_labels()
        ax.legend(lines1 + lines2 + lines3, labels1 + labels2 + labels3, 
                  loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2, frameon=True, fontsize=12)
        
        self._apply_premium_style(ax)
        for spine in ax2.spines.values():
            spine.set_visible(False)
        for spine in ['top', 'right', 'bottom']:
            ax3.spines[spine].set_visible(False)
        ax3.spines['left'].set_color('#27AE60')
        ax3.tick_params(axis='y', colors='#27AE60')
            
        plt.tight_layout()
        fig.subplots_adjust(bottom=0.2, left=0.15) # More space for the extra left axis
        self._add_watermark(fig)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, f'{self.scenario_name}_Interest_Paid.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        # Store data for Excel
        self.charts_data.append(("Bank Loans: Interest Paid", df_plot[['Interest_Paid (M€)']]))

    def _plot_prices(self):
        """Plots all prices used in the simulation (EUA and Resources)."""
        price_series = {}
        
        # 1. Carbon Prices (EUA)
        if self.data.time_series.carbon_prices:
            price_series['EUA'] = {
                'data': self.data.time_series.carbon_prices,
                'name': 'EUA (Carbon Price)',
                'unit': '€/tCO2',
                'color': '#2C3E50'
            }
            
        # 2. Resource Prices
        colors = ['#2980B9', '#8E44AD', '#D35400', '#C0392B', '#16A085', '#273746', '#F39C12', '#BDC3C7']
        color_idx = 0
        for r_id, p_dict in self.data.time_series.resource_prices.items():
            if p_dict:
                res = self.data.resources.get(r_id)
                name = res.name if res and res.name else r_id
                unit = res.unit if res and res.unit else 'unit'
                price_series[r_id] = {
                    'data': p_dict,
                    'name': name,
                    'unit': f"€/{unit}",
                    'color': colors[color_idx % len(colors)]
                }
                color_idx += 1
                
        if not price_series:
            return
            
        n_plots = len(price_series)
        import math
        n_cols = math.ceil(math.sqrt(n_plots))
        n_rows = math.ceil(n_plots / n_cols)
        
        # Adjust figsize based on columns and rows
        # We want to keep it somewhat squared, so maybe 4-5 inches per subplot
        fig, axes = plt.subplots(n_rows, n_cols, figsize=(4 * n_cols, 3 * n_rows), facecolor='white')
        
        if n_plots == 1:
            axes = np.array([axes])
        axes = axes.flatten()
        
        i = -1
        for i, (key, info) in enumerate(price_series.items()):
            ax = axes[i]
            years = sorted(info['data'].keys())
            values = [info['data'][y] for y in years]
            
            p_ln = ax.plot(years, values, color=info['color'], linewidth=2.5, marker='o', 
                    markersize=4, markerfacecolor='white', markeredgewidth=1.5, label='Price')
            
            ax.set_title(info['name'].upper(), fontsize=12, weight='bold', pad=10)
            ax.set_ylabel(info['unit'], fontsize=10, weight='semibold', color=info['color'])
            ax.set_xlabel('Year', fontsize=9)
            
            # --- CO2 Emissions Comparison ---
            handles, labels = ax.get_legend_handles_labels()
            
            if key in self.data.time_series.other_emissions_factors:
                em_data = self.data.time_series.other_emissions_factors[key]
                em_values = [em_data.get(y, 0.0) for y in years]
                
                ax2 = ax.twinx()
                em_ln = ax2.plot(years, em_values, color='#27AE60', linewidth=2, linestyle='--', 
                        marker='s', markersize=3, alpha=0.7, label='CO2 intensity')
                
                ax2.set_ylabel('tCO2 / unit', fontsize=9, color='#27AE60', weight='semibold')
                ax2.tick_params(axis='y', labelcolor='#27AE60')
                
                max_em = max(em_values) if em_values and max(em_values) > 0 else 1.0
                ax2.set_ylim(bottom=0, top=max_em * 1.6)
                
                # Combine handles for legend
                h2, l2 = ax2.get_legend_handles_labels()
                handles += h2
                labels += l2
                
                # Apply premium style to ax2 spines
                for spine in ['top', 'left', 'bottom']:
                    ax2.spines[spine].set_visible(False)
                ax2.spines['right'].set_color('#27AE60')
                ax2.spines['right'].set_alpha(0.5)

            # --- Formatting ---
            ax.xaxis.set_major_locator(ticker.MultipleLocator(5))
            ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x:,.2f}'))
            
            # Ensure axes start at 0 and apply scaling factors
            max_price = max(values) if values else 1.0
            ax.set_ylim(bottom=0, top=max_price * 1.2)
            
            if key in self.data.time_series.other_emissions_factors:
                # ax2 was created in the CO2 block above, but we need to reference it here or handle it there.
                # Actually, it's cleaner to handle ax2 limit inside the CO2 block.
                pass
            
            self._apply_premium_style(ax)
            
            # Add Legend if CO2 is present or just to be safe
            if len(handles) > 1:
                ax.legend(handles, labels, loc='best', fontsize=12, frameon=True, framealpha=0.8)

            
        # Hide unused axes
        for j in range(i + 1, len(axes)):
            axes[j].axis('off')
            
        fig.suptitle('SIMULATION PRICE PARAMETERS', fontsize=18, weight='bold', y=0.98)
        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        self._add_watermark(fig)
        self._add_scenario_label(fig)
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, f'{self.scenario_name}_Simulation_Prices.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        # Store for Excel
        df_prices = pd.DataFrame({info['name']: info['data'] for info in price_series.values()}).sort_index()
        df_prices.index.name = 'Year'
        self.charts_data.append(("Simulation Prices", df_prices))

    def _plot_co2_abatement_cost(self, df: pd.DataFrame):
        """Plot the Marginal Abatement Cost (MAC) per technology implementation."""
        # --- Corporate & Premium Design ---
        fig, ax = plt.subplots(figsize=(14, 10))
        fig.set_facecolor('#F8F9FA')
        ax.set_facecolor('#F8F9FA')
        
        # Sort values
        df_plot = df.sort_values(by='MAC (€/tCO2)')
        
        # Create premium colors - Palette based on cost (green to red)
        # Handle zero or negative costs explicitly
        norm = plt.Normalize(df_plot['MAC (€/tCO2)'].min(), df_plot['MAC (€/tCO2)'].max())
        import matplotlib.cm as cm
        colors = cm.RdYlGn_r(norm(df_plot['MAC (€/tCO2)'].values))
        
        # Adjust colors for negative MAC (profitable)
        for i, (idx, row) in enumerate(df_plot.iterrows()):
            if row['MAC (€/tCO2)'] < 0:
                colors[i] = [0.1, 0.6, 0.2, 0.9] # Solid green for profitable
        
        # Visual distinction for Potential vs Invested
        alphas = [0.9 if s == 'Invested' else 0.45 for s in df_plot['Status']]
        edge_colors = ['#2C3E50' if s == 'Invested' else '#7F8C8D' for s in df_plot['Status']]
        linestyles = ['-' if s == 'Invested' else '--' for s in df_plot['Status']]
        
        x_labels = df_plot['Display Label'] if 'Display Label' in df_plot.columns else df_plot['Project']
        y_capex = df_plot['MAC CAPEX (€/tCO2)'].values
        y_opex = df_plot['MAC OPEX (€/tCO2)'].values
        y_total = df_plot['MAC (€/tCO2)'].values

        # Plot bars item by item to handle per-technology alpha
        for i in range(len(x_labels)):
            label = x_labels.iloc[i] if hasattr(x_labels, 'iloc') else x_labels[i]
            
            # Plot CAPEX Part
            ax.bar(label, y_capex[i], color='#3498DB', alpha=alphas[i], 
                     edgecolor=edge_colors[i], linewidth=1.2, width=0.65, 
                     linestyle=linestyles[i], label='CAPEX (Investment) Component' if i == 0 else "", zorder=3)
            
            # Plot OPEX Part (Stacked)
            ax.bar(label, y_opex[i], bottom=y_capex[i], color='#E67E22', alpha=alphas[i],
                      edgecolor=edge_colors[i], linewidth=1.2, width=0.65,
                      linestyle=linestyles[i], label='OPEX (Operational) Component' if i == 0 else "", zorder=3)
        
        # Add numeric labels at the top of the TOTAL bar
        for i, (label, val) in enumerate(zip(x_labels, y_total)):
            va = 'bottom' if val >= 0 else 'top'
            offset = 5 if val >= 0 else -15
            ax.annotate(f'{val:,.0f} €/t',
                        xy=(i, val),
                        xytext=(0, offset),
                        textcoords="offset points",
                        ha='center', va=va, fontsize=10, weight='bold', color='#2C3E50')

        # Average CO2 Price Line
        avg_co2_price = sum(self.data.time_series.carbon_prices.values()) / len(self.data.time_series.carbon_prices) if self.data.time_series.carbon_prices else 0.0
        if avg_co2_price > 0:
            ax.axhline(avg_co2_price, color='#E74C3C', linestyle='--', linewidth=2, label=f'Avg Carbon Price ({avg_co2_price:,.0f} €/t)', zorder=4)
            
            # Fill region below carbon price as "Economic Zone"
            ax.fill_between([-1, len(df_plot)], 0, avg_co2_price, color='#2ECC71', alpha=0.07, label='Profitable Zone (Cost < Tax)')

        # Style refinement
        plt.title('CO2 ABATEMENT COST BY TECHNOLOGY (MAC)', fontsize=18, weight='bold', pad=30, color='#1A1A1A')
        plt.ylabel('Cost of Abatement (€ / tCO2 avoided)', fontsize=13, weight='semibold', color='#333333')
        plt.xlabel('Implemented Technology per Process', fontsize=13, weight='semibold', color='#333333')
        
        plt.xticks(rotation=45, ha='right', fontsize=11)
        ax.grid(axis='y', linestyle=':', alpha=0.6, zorder=0)
        
        # Zero line
        ax.axhline(0, color='black', linewidth=1.0, zorder=2)
        
        # Legend
        ax.legend(loc='best', fontsize=12, frameon=True, facecolor='white', framealpha=0.9)
        
        # Add info box about total abated
        total_tons = df_plot['Total Abated (tCO2)'].sum()
        ax.text(0.98, 0.02, f"Total Simulation Abatement: {total_tons/1000:,.0f} ktCO2", 
                transform=ax.transAxes, ha='right', va='bottom', 
                bbox=dict(facecolor='white', alpha=0.8, edgecolor='#CCCCCC'), fontsize=11, weight='bold')

        self._apply_premium_style(ax)
        plt.tight_layout()
        self._add_watermark(fig)
        self._add_scenario_label(fig)
        
        os.makedirs(self.results_dir, exist_ok=True)
        plt.savefig(os.path.join(self.results_dir, f'{self.scenario_name}_CO2_Abatement_Cost.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        # Store data for Excel
        self.charts_data.append(("CO2 Abatement Cost (MAC)", df_plot))
