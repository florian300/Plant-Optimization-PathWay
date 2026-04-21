import pandas as pd
import numpy as np
import warnings
from rich import print
from tqdm import tqdm
from typing import Dict, Any, Tuple, List
from .model import Parameters, Resource, Technology, TimeSeriesData, EntityState, PathFinderData, Objective, Process, GrantParams, CCfDParams, BankLoan, DACParams, CreditParams, ReportingToggles, SensitivityParams

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
pd.set_option('future.no_silent_downcasting', True)

class PathFinderParser:
    def __init__(self, file_path: str, verbose: bool = False):
        self.file_path = file_path
        self.xl = pd.ExcelFile(file_path)
        self.verbose = verbose
        self.sim_row_found = False        # Track if SIM row exists in OverView
        self.all_scenarios_meta = []      # Store all scenarios found before filtering

    def _parse_scenarios(self) -> list:
        """Parse the MODELING START/END block from OverView and return list of {id, name} dicts.

        SIM row behaviour:
            - SIM row absent          → simulate all SC-DES scenarios (backward-compatible default)
            - SIM row present + IDs   → simulate only the listed scenario IDs
            - SIM row present + empty → simulate nothing (user explicitly left it blank)
        """
        df_overview = self.xl.parse('OverView', header=None)
        all_scenarios = []
        to_simulate = []
        self.sim_row_found = False
        in_modeling = False
        for _, row in df_overview.iterrows():
            vals_upper = [str(x).strip().upper() for x in row.values if pd.notna(x) and str(x).strip()]
            if 'MODELING' in vals_upper and 'START' in vals_upper:
                in_modeling = True
                continue
            if 'MODELING' in vals_upper and 'END' in vals_upper:
                break
            if in_modeling:
                raw = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                if not raw or raw[0].startswith('**'):
                    continue

                if 'SC-DES' in vals_upper:
                    # raw[0] is 'SC-DES', raw[1] is name, raw[2] is ID
                    if len(raw) >= 3:
                        sc_name = raw[1]
                        sc_id   = raw[2]
                        all_scenarios.append({'id': sc_id, 'name': sc_name})

                elif 'SIM' in vals_upper:
                    self.sim_row_found = True   # Mark that a SIM row was explicitly defined
                    # Collect all tokens after SIM as scenario IDs to simulate
                    if len(raw) > 1:
                        for token in raw[1:]:
                            # Skip common header placeholders if present
                            if token.upper() not in ['SC1', 'SC2', 'SC3', 'SC4', 'SC...', 'SC…']:
                                to_simulate.append(token.upper())

        self.all_scenarios_meta = [s.copy() for s in all_scenarios]

        # SIM row was present but left empty → user wants no simulations
        if self.sim_row_found and not to_simulate:
            return []

        # SIM row had IDs → only run those
        if to_simulate:
            filtered = [s for s in all_scenarios if s['id'].upper() in to_simulate]
            return filtered if filtered else all_scenarios

        # SIM row was absent entirely → backward-compatible: run all SC-DES scenarios
        return all_scenarios

    def _find_blocks(self, df: pd.DataFrame) -> list:
        """Find START and END blocks in a dataframe to isolate tables."""
        blocks = []
        for i, row in df.iterrows(): # tqdm removed as it is too noisy in sensitivity runs
            for j, val in enumerate(row):
                if isinstance(val, str):
                    clean_val = val.strip().upper()
                    if clean_val in ['START', 'END']:
                        prefix = [str(x).strip() for x in row[:j] if pd.notna(x) and str(x).strip() != '']
                        prefix_str = " ".join(prefix).strip()
                        blocks.append({
                            'row': i,
                            'col': j,
                            'type': clean_val,
                            'prefix': prefix_str
                        })
        return blocks

    def _extract_block_data(self, df: pd.DataFrame, start_row: int, end_row: int) -> pd.DataFrame:
        """Extracts data between START and END markers, assuming the first row after START is the header."""
        if end_row <= start_row + 1:
            return pd.DataFrame()
        
        # Row immediately after START is usually the header, but might be empty. Let's find the first non-completely-empty row.
        block_df = df.iloc[start_row+1:end_row].dropna(how='all')
        if block_df.empty:
            return block_df
            
        block_df.columns = block_df.iloc[0]
        block_df = block_df.iloc[1:].reset_index(drop=True)
        # remove columns that are entirely NaN
        block_df = block_df.dropna(axis=1, how='all')
        return block_df

    def _interpolate_linear(self, series: pd.Series) -> pd.Series:
        """Replace 'LINEAR INTER' or keywords containing 'BROWNIEN' with NaN and interpolate linearly.
        If 'BROWNIEN' was detected, add a Brownian Bridge with high amplitude (~25%)."""
        
        # Robust detection: case-insensitive, check if 'BROWNIEN' is in the string
        str_series = series.astype(str).str.strip().str.upper()
        is_brownian = str_series.str.contains('BROWNIEN').any()
        
        # Identify keywords to replace
        # We replace known markers and anything containing BROWNIEN
        to_replace = ['LINEAR INTER', 'LINEAR INTER ']
        # Add actual values that contain BROWNIEN
        to_replace.extend(series[str_series.str.contains('BROWNIEN')].unique())
        
        res = series.replace(to_replace, np.nan)
        res = pd.to_numeric(res, errors='coerce')
        
        # Identify blocks of NaNs to interpolate
        mask = res.isna()
        if not mask.any():
            return res
            
        res = res.interpolate(method='linear')
        res = res.bfill().ffill().fillna(0.0)

        if is_brownian:
            vals = res.values
            n = len(vals)
            noise = np.zeros(n)
            
            # Identify contiguous NaN segments using positional indexing
            diff = np.diff(mask.values.astype(int), prepend=0)
            starts = np.where(diff == 1)[0]
            ends = np.where(diff == -1)[0]
            
            if len(starts) > len(ends):
                ends = np.append(ends, n)
            
            for s_idx, e_idx in zip(starts, ends):
                start_pos = s_idx - 1
                end_pos = e_idx
                
                if start_pos >= 0 and end_pos < n:
                    segment_len = end_pos - start_pos
                    t_axis = np.arange(segment_len + 1)
                    avg_val = (vals[start_pos] + vals[end_pos]) / 2.0
                    
                    # Amplitude set to 12.5% of the average value
                    # Symmetrical distribution (mean 0) ensures up and down variations
                    sigma = 0.125 * abs(avg_val) if avg_val != 0 else 1.0 
                    
                    steps = np.random.normal(0, sigma, segment_len)
                    w = np.cumsum(np.insert(steps, 0, 0))
                    bridge = w - (t_axis / segment_len) * w[-1]
                    
                    noise[start_pos : end_pos + 1] += bridge
            
            res = pd.Series(vals + noise, index=res.index)
            
        return res

    def _normalize_token(self, raw_val) -> str:
        import pandas as pd
        return str(raw_val).strip().upper() if pd.notna(raw_val) else ""

    def _is_emission_type(self, resource_type: str) -> bool:
        return 'EMISS' in self._normalize_token(resource_type)

    def _is_primary_emission_resource(self, resource_obj) -> bool:
        name_upper = self._normalize_token(resource_obj.name)
        return self._is_emission_type(resource_obj.type) and ('CO2' in name_upper or 'CARBON' in name_upper)

    def _parse_overview_settings(self, df_overview, blocks_overview):
        import pandas as pd
        from .model import ReportingToggles
        reporting_toggles = ReportingToggles()
        charts_start = next((b['row'] for b in blocks_overview if b['type'] == 'START' and b['prefix'] == 'CHARTS'), None)
        charts_end = next((b['row'] for b in blocks_overview if b['type'] == 'END' and b['prefix'] == 'CHARTS'), None)
        
        if charts_start is not None and charts_end is not None:
            df_charts = self._extract_block_data(df_overview, charts_start, charts_end)
            if not df_charts.empty:
                for _, row in df_charts.iterrows():
                    row_vals = [str(x).strip().upper() for x in row.values if pd.notna(x) and str(x).strip()]
                    if not row_vals or row_vals[0].startswith('**'):
                        continue
                    name = row_vals[0]
                    if 'YES' in row_vals or 'NO' in row_vals:
                        is_yes = 'YES' in row_vals
                        if "EXCEL DATA" in name: reporting_toggles.results_excel = is_yes
                        elif "ENERGY MIX" in name: reporting_toggles.chart_energy_mix = is_yes
                        elif "CO2 TRAJECTORY" in name: reporting_toggles.chart_co2_trajectory = is_yes
                        elif "INDIRECT EMISSIONS" in name: reporting_toggles.chart_indirect_emissions = is_yes
                        elif "INVESTMENT PLAN" in name or "INVESTMENT COSTS" in name: reporting_toggles.chart_investment_costs = is_yes
                        elif "RESSOURCES OPEX" in name or "RESOURCE OPEX" in name or "TOTAL OPEX" in name or "CCS OPEX" in name: reporting_toggles.chart_total_opex = is_yes
                        elif "CARBON TAX" in name: reporting_toggles.chart_carbon_tax_avoided = is_yes
                        elif "FINANCING" in name: reporting_toggles.chart_external_financing = is_yes
                        elif "TRANSITION COST" in name: reporting_toggles.chart_transition_costs = is_yes
                        elif "CARBON PRICE" in name: reporting_toggles.chart_carbon_prices = is_yes
                        elif "SIMULATION PRICES" in name or "RESOURCE PRICES" in name: reporting_toggles.chart_resource_prices = is_yes
                        elif "INTEREST PAID" in name: reporting_toggles.chart_interest_paid = is_yes
                        elif "ABATEMENT" in name or "MAC CURVE" in name: reporting_toggles.chart_co2_abatement_cost = is_yes
                    elif "CAP" in name:
                        for val in row_vals[1:]:
                            try:
                                clean_val = val.replace('M€', '').replace('M', '').replace(' ', '').replace(',', '.')
                                reporting_toggles.investment_cap = float(clean_val)
                                break
                            except ValueError:
                                continue
        
        start_year, duration, time_limit, mip_gap, relax_integrality, discount_rate, run_project = 2025, 25, 60.0, 0.90, False, 0.0, True
        for i, row in df_overview.iterrows():
            row_vals = [str(x).strip().upper() if pd.notna(x) else "" for x in row]
            if "YEAR START" in row_vals:
                idx = row_vals.index("YEAR START")
                if len(row) > idx + 1:
                    start_year = int(row.iloc[idx + 1])
            if "SIMULATION TIME (IN YEAR)" in row_vals:
                idx = row_vals.index("SIMULATION TIME (IN YEAR)")
                if len(row) > idx + 1:
                    duration = int(row.iloc[idx + 1])
            if "DURURATION SIMULATION (S)" in row_vals or "DURATION SIMULATION (S)" in row_vals:
                keyword = "DURURATION SIMULATION (S)" if "DURURATION SIMULATION (S)" in row_vals else "DURATION SIMULATION (S)"
                idx = row_vals.index(keyword)
                if len(row) > idx + 1:
                    try: time_limit = float(row.iloc[idx + 1])
                    except ValueError: pass
            if "RELAX INTEGRAL" in row_vals:
                idx = row_vals.index("RELAX INTEGRAL")
                if len(row) > idx + 1:
                    val = str(row.iloc[idx + 1]).strip().upper()
                    relax_integrality = (val == 'YES')
            if "ERROR SIMULATION (%)" in row_vals:
                idx = row_vals.index("ERROR SIMULATION (%)")
                if len(row) > idx + 1:
                    raw_val = str(row.iloc[idx + 1]).strip()
                    try:
                        val = float(raw_val.replace('%', ''))
                        if '%' in raw_val or val > 1.0: val /= 100.0
                        mip_gap = val
                    except ValueError: pass
            if "DISCOUNT RATE (%)" in row_vals or "DISCOUNT RATE" in row_vals:
                keyword = "DISCOUNT RATE (%)" if "DISCOUNT RATE (%)" in row_vals else "DISCOUNT RATE"
                idx = row_vals.index(keyword)
                if len(row) > idx + 1:
                    raw_val = str(row.iloc[idx + 1]).strip()
                    try:
                        val = float(raw_val.replace('%', ''))
                        if '%' in raw_val or val > 1.0: val /= 100.0
                        discount_rate = val
                    except ValueError: pass
            if "RUN PROJECT ?" in row_vals or "RUN PROJECT ? (YES/NO)" in row_vals:
                keyword = "RUN PROJECT ?" if "RUN PROJECT ?" in row_vals else "RUN PROJECT ? (YES/NO)"
                idx = row_vals.index(keyword)
                if len(row) > idx + 1:
                    val = str(row.iloc[idx + 1]).strip().upper()
                    run_project = (val == 'YES')
        return reporting_toggles, start_year, duration, time_limit, mip_gap, relax_integrality, discount_rate, run_project

    def _parse_entities_cluster(self, df_overview, blocks_overview):
        import pandas as pd
        cluster_start = next((b['row'] for b in blocks_overview if b['type'] == 'START' and b['prefix'] == 'CLUSTER'), None)
        cluster_end = next((b['row'] for b in blocks_overview if b['type'] == 'END' and b['prefix'] == 'CLUSTER'), None)
        entities_info = {}
        entities = []
        if cluster_start is not None and cluster_end is not None:
            df_cluster = self._extract_block_data(df_overview, cluster_start, cluster_end)
            cluster_col_map = {self._normalize_token(c): c for c in df_cluster.columns}
            id_col = cluster_col_map.get('ID')
            name_col = cluster_col_map.get('NAME')
            prod_col = cluster_col_map.get('PRODUCTION')
            sheet_col = cluster_col_map.get('SHEET')
            if id_col is not None:
                for _, row in df_cluster.iterrows():
                    e_id = str(row.get(id_col, '')).strip()
                    if e_id and e_id.lower() != 'nan':
                        prod = row.get(prod_col, 0.0) if prod_col is not None else 0.0
                        try: prod = float(prod)
                        except Exception: prod = 0.0
                        sheet_name = str(row.get(sheet_col, '')).strip() if sheet_col is not None else ''
                        entity_name = str(row.get(name_col, '')).strip() if name_col is not None else str(e_id)
                        if not entity_name or entity_name.lower() == 'nan': entity_name = e_id
                        entities_info[e_id] = {'production': prod, 'sheet': sheet_name, 'name': entity_name}
                entities = list(entities_info.keys())
        return entities_info, entities

    def _parse_resources(self, df_overview, blocks_overview):
        import pandas as pd
        from .model import Resource
        data_start = next((b['row'] for b in blocks_overview if b['type'] == 'START' and b['prefix'] == 'DATA'), None)
        data_end = next((b['row'] for b in blocks_overview if b['type'] == 'END' and b['prefix'] == 'DATA'), None)
        resources_dict = {}
        if data_start is not None and data_end is not None:
            df_data = self._extract_block_data(df_overview, data_start, data_end)
            for idx, row in df_data.iterrows():
                row_list = row.tolist()
                if 'ID' in row_list:
                    df_data.columns = row_list
                    df_data = df_data.iloc[idx+1:]
                    break
            if 'ID' in df_data.columns and 'TYPE' in df_data.columns and 'UNIT' in df_data.columns:
                if 'NAME' not in df_data.columns:
                    raise ValueError("DATA block is missing required NAME column for resources")
                for _, row in df_data.iterrows():
                    res_id = str(row['ID']).strip()
                    if res_id and pd.notna(res_id) and res_id != 'nan':
                        res_name = str(row.get('NAME', '')).strip()
                        if not res_name or res_name.lower() == 'nan':
                            raise ValueError(f"Resource '{res_id}' is missing NAME in DATA block")
                        category = str(row.get('CATEGORY', 'Other')).strip()
                        if not category or category.lower() == 'nan': category = 'Other'
                        resource_type = str(row.get('RESSOURCE TYPE', 'GENERIC')).strip().upper()
                        if not resource_type or resource_type.lower() == 'NAN': resource_type = 'GENERIC'
                        resources_dict[res_id] = Resource(id=res_id, type=str(row['TYPE']), unit=str(row['UNIT']), name=res_name, category=category, resource_type=resource_type)

        purchases_start = next((b['row'] for b in blocks_overview if b['type'] == 'START' and 'PURCHASES' in self._normalize_token(b['prefix'])), None)
        purchases_end = next((b['row'] for b in blocks_overview if b['type'] == 'END' and 'PURCHASES' in self._normalize_token(b['prefix'])), None)
        if purchases_start is not None and purchases_end is not None:
            df_purchases = self._extract_block_data(df_overview, purchases_start, purchases_end)
            if not df_purchases.empty:
                col_map = {self._normalize_token(c): c for c in df_purchases.columns}
                id_col = col_map.get('RESSOURCE ID') or col_map.get('RESOURCE ID') or col_map.get('ID')
                buy_sell_col = col_map.get('BUY/SELL') or col_map.get('TYPE')
                if id_col is not None and buy_sell_col is not None:
                    for _, row in df_purchases.iterrows():
                        res_id = str(row.get(id_col, '')).strip()
                        bs_val = str(row.get(buy_sell_col, '')).strip().upper()
                        if res_id and res_id in resources_dict:
                            if 'BUY' in bs_val and 'SELL' in bs_val:
                                resources_dict[res_id].can_buy = True
                                resources_dict[res_id].can_sell = True
                            elif 'BOTH' in bs_val:
                                resources_dict[res_id].can_buy = True
                                resources_dict[res_id].can_sell = True
                            elif 'BUY' in bs_val:
                                resources_dict[res_id].can_buy = True
                            elif 'SELL' in bs_val:
                                resources_dict[res_id].can_sell = True
        return resources_dict

    def _parse_unit_conversions(self, df_overview, blocks_overview):
        import pandas as pd
        unit_conversions = {}
        unit_conv_start = next((b['row'] for b in blocks_overview if b['type'] == 'START' and 'UNIT CONVERSION' in self._normalize_token(b['prefix'])), None)
        unit_conv_end = next((b['row'] for b in blocks_overview if b['type'] == 'END' and 'UNIT CONVERSION' in self._normalize_token(b['prefix'])), None)
        if unit_conv_start is not None and unit_conv_end is not None:
            df_unit_conv = self._extract_block_data(df_overview, unit_conv_start, unit_conv_end)
            if not df_unit_conv.empty:
                col_map = {self._normalize_token(c): c for c in df_unit_conv.columns}
                unit_in_col = col_map.get('UNIT IN')
                unit_out_col = col_map.get('UNIT OUT')
                factor_col = col_map.get('FACTOR')
                if unit_in_col is None or unit_out_col is None or factor_col is None:
                    raise ValueError("UNIT CONVERSIONS block must contain UNIT IN, UNIT OUT, and FACTOR columns")
                for row_idx, row in df_unit_conv.iterrows():
                    unit_in = self._normalize_token(row.get(unit_in_col, ''))
                    unit_out = self._normalize_token(row.get(unit_out_col, ''))
                    factor_raw = row.get(factor_col, None)
                    if not unit_in or not unit_out or unit_in == 'NAN' or unit_out == 'NAN': continue
                    try: factor = float(str(factor_raw).replace(',', '.').strip())
                    except Exception: raise ValueError(f"Invalid FACTOR in UNIT CONVERSIONS at row {row_idx}: {factor_raw}")
                    if factor == 0.0: raise ValueError(f"UNIT CONVERSIONS factor cannot be zero for {unit_in} -> {unit_out}")
                    unit_conversions[(unit_in, unit_out)] = factor
                    unit_conversions[(unit_out, unit_in)] = 1.0 / factor
        return unit_conversions

    def _parse_objectives(self, df_overview, blocks_overview):
        import pandas as pd
        from .model import Objective
        objectives_list = []
        obj_start = next((b['row'] for b in blocks_overview if b['type'] == 'START' and b['prefix'] == 'OBJECTIVES'), None)
        obj_end = next((b['row'] for b in blocks_overview if b['type'] == 'END' and b['prefix'] == 'OBJECTIVES'), None)
        if obj_start is not None and obj_end is not None:
            df_obj = self._extract_block_data(df_overview, obj_start, obj_end)
            if not df_obj.empty:
                col_mapping = {}
                for c in df_obj.columns:
                    c_str = str(c).upper()
                    if 'LIMIT' in c_str: col_mapping[c] = 'LIMIT'
                    elif 'YEAR AT WHICH' in c_str: col_mapping[c] = 'TARGET_YEAR'
                    elif 'VALUE' in c_str: col_mapping[c] = 'CAP_VALUE'
                    elif 'COMPARAISON' in c_str: col_mapping[c] = 'COMP_YEAR'
                    elif 'ENTITY' in c_str: col_mapping[c] = 'ENTITY'
                    elif 'RESSOURCE' in c_str: col_mapping[c] = 'RESOURCE'
                    elif 'INTERPOLATION' in c_str: col_mapping[c] = 'INTERPOLATION'
                    elif 'GROUP' in c_str: col_mapping[c] = 'GROUP'
                    elif 'NAME' in c_str: col_mapping[c] = 'NAME'
                    elif 'PENALTY' in c_str or 'PENALITY' in c_str: col_mapping[c] = 'PENALTY'
                df_obj = df_obj.rename(columns=col_mapping)
                for _, row in df_obj.iterrows():
                    res = str(row.get('RESOURCE', '')).strip()
                    if res and res != 'nan':
                        lim = str(row.get('LIMIT', 'CAP')).strip().upper()
                        if 'CAP' in lim: lim = 'CAP'
                        elif 'MIN' in lim: lim = 'MIN'
                        elif 'MAX' in lim: lim = 'MAX'
                        else: lim = 'CAP'
                        try: t_year = int(row.get('TARGET_YEAR', 0))
                        except: t_year = 0
                        try: 
                            raw_cap = str(row.get('CAP_VALUE', 0.0))
                            has_pct = '%' in raw_cap
                            c_val = float(raw_cap.replace('%','').replace(',','.'))
                            if has_pct: c_val /= 100.0
                        except: c_val = 0.0
                        c_year = None
                        try: 
                            if pd.notna(row.get('COMP_YEAR')): c_year = int(row.get('COMP_YEAR', 0))
                        except: pass
                        ent = str(row.get('ENTITY', 'ALL')).strip()
                        mode = str(row.get('INTERPOLATION', 'NONE')).strip().upper()
                        if mode not in ['LINEAR', 'NONE']: mode = 'NONE'
                        group = str(row.get('GROUP', '')).strip()
                        penalty = str(row.get('PENALTY', 'AT ALL COST')).strip().upper()
                        if not penalty or penalty == 'NAN': penalty = 'AT ALL COST'
                        if penalty in ['PENALTIES', 'PENALITIES']: penalty = 'PENALTIES'
                        objectives_list.append(Objective(
                            entity=ent, resource=res, limit_type=lim, target_year=t_year, cap_value=c_val,
                            comparison_year=c_year, mode=mode, group=group, name=str(row.get('NAME', '')).strip(), penalty_type=penalty
                        ))
        return objectives_list

    def _parse_technologies(self, df_tech, scenario_id, active_sc_name, resources_dict, years_list):
        import pandas as pd
        from .model import Technology
        blocks_tech = self._find_blocks(df_tech)
        technologies_dict = {}
        tecs_start = next((b['row'] for b in blocks_tech if b['type'] == 'START' and b['prefix'] == 'TECS'), None)
        tecs_end = next((b['row'] for b in blocks_tech if b['type'] == 'END' and b['prefix'] == 'TECS'), None)
        if tecs_start is not None and tecs_end is not None:
            df_tecs = self._extract_block_data(df_tech, tecs_start, tecs_end)
            for idx, row in df_tecs.iterrows():
                row_list = row.tolist()
                if 'ID' in row_list and 'NAME' in row_list:
                    df_tecs.columns = row_list
                    df_tecs = df_tecs.iloc[idx+1:]
                    break
            if 'NAME' not in df_tecs.columns: raise ValueError("TECS block is missing required NAME column for technologies")
            tecs_col_map = {str(c).strip().upper(): c for c in df_tecs.columns}
            is_ci_col = next((tecs_col_map[k] for k in tecs_col_map if k in ["IS CONTINUOUS IMPROVEMENT", "CONTINUOUS IMPROVEMENT", "IS_CI"]), None)
            tech_category_col = next((tecs_col_map[k] for k in tecs_col_map if k in ["TECH CATEGORY", "TECHNOLOGY CATEGORY", "CATEGORY"]), None)
            def _parse_bool(raw_val):
                v = str(raw_val).strip().upper()
                if v in ['YES', 'TRUE', '1', 'Y']: return True
                if v in ['NO', 'FALSE', '0', 'N', '', 'NAN']: return False
                try: return float(v) != 0.0
                except Exception: return False
            for _, row in df_tecs.iterrows():
                t_id = str(row.get('ID', '')).strip()
                if t_id and t_id != 'nan':
                    imp_time_raw = row.get('IMPLEMANTATION TIME (YEAR)') or row.get('IMPLEMENTATION TIME (YEAR)')
                    imp_time = int(imp_time_raw) if pd.notna(imp_time_raw) else 1
                    t_name = str(row.get('NAME', '')).strip()
                    if not t_name or t_name.lower() == 'nan': raise ValueError(f"Technology '{t_id}' is missing NAME in TECS block")
                    is_continuous_improvement = False
                    if is_ci_col is not None: is_continuous_improvement = _parse_bool(row.get(is_ci_col, False))
                    if t_id.upper() == 'UP': is_continuous_improvement = True
                    tech_category = "Standard"
                    if tech_category_col is not None:
                        raw_cat = str(row.get(tech_category_col, '')).strip()
                        if raw_cat and raw_cat.lower() != 'nan': tech_category = raw_cat
                    t_id_upper = t_id.upper()
                    if tech_category == 'Standard' and ('CCS' in t_id_upper or 'CCU' in t_id_upper):
                        tech_category = 'Carbon Capture'
                    technologies_dict[t_id] = Technology(id=t_id, name=t_name, implementation_time=imp_time, capex=0.0, opex=0.0, impacts={}, is_continuous_improvement=is_continuous_improvement, tech_category=tech_category)
        
        euro_start = next((b['row'] for b in blocks_tech if b['type'] == 'START' and ('SPECS' in b['prefix'] or 'COMPATIBILITIES' in b['prefix']) and 'TECHNICAL' not in b['prefix']), None)
        euro_end = next((b['row'] for b in blocks_tech if b['type'] == 'END' and ('SPECS' in b['prefix'] or 'COMPATIBILITIES' in b['prefix']) and 'TECHNICAL' not in b['prefix']), None)
        raw_tech_capex = {t: {} for t in technologies_dict}
        raw_tech_opex = {t: {} for t in technologies_dict}
        raw_tech_capex_links = {}
        raw_tech_opex_links = {}
        if euro_start is not None and euro_end is not None:
            df_euros = self._extract_block_data(df_tech, euro_start, euro_end)
            for idx, row in df_euros.iterrows():
                row_list = [str(x).upper().strip() for x in row.tolist()]
                if 'TECH ID' in row_list and 'COST' in row_list:
                    df_euros.columns = row_list
                    df_euros = df_euros.iloc[idx+1:]
                    break
            for _, row in df_euros.iterrows():
                t_id = str(row.get('TECH ID', '')).strip()
                if t_id in technologies_dict:
                    row_scenario = str(row.get('SCENARIO', '')).strip().upper()
                    if scenario_id and row_scenario and row_scenario not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]:
                        continue
                    cost_nature = str(row.get('TYPE (CAPEX/OPEX)', '')).strip().upper()
                    if not cost_nature or cost_nature == 'NAN':
                        exp_type = str(row.get('TYPE (VARIABLE/FIXED)', '')).strip().upper()
                        cost_nature = 'CAPEX' if exp_type == 'FIXED' else 'OPEX' if exp_type == 'VARIABLE' else ''
                    try: cost_val = float(row.get('COST', 0))
                    except: cost_val = 0.0
                    y_str = str(row.get('YEAR', 'ALL')).strip().upper()
                    try: year_val = int(float(y_str))
                    except: year_val = y_str
                    per_unit_str = str(row.get('PER UNIT ?', 'NO')).strip().upper()
                    is_per_unit = per_unit_str == 'YES'
                    unit_str = str(row.get('UNIT', '')).strip()
                    if cost_nature == 'CAPEX':
                        technologies_dict[t_id].capex_per_unit = is_per_unit
                        technologies_dict[t_id].capex_unit = unit_str
                        if year_val == 'ALL':
                            if technologies_dict[t_id].capex == 0.0: technologies_dict[t_id].capex = cost_val
                        elif isinstance(year_val, str): raw_tech_capex_links[t_id] = (year_val, cost_val)
                        else: raw_tech_capex[t_id][year_val] = cost_val
                    elif cost_nature == 'OPEX':
                        technologies_dict[t_id].opex_per_unit = is_per_unit
                        technologies_dict[t_id].opex_unit = unit_str
                        if year_val == 'ALL':
                            if technologies_dict[t_id].opex == 0.0: technologies_dict[t_id].opex = cost_val
                        elif isinstance(year_val, str): raw_tech_opex_links[t_id] = (year_val, cost_val)
                        else: raw_tech_opex[t_id][year_val] = cost_val
            for t_id in technologies_dict:
                if raw_tech_capex[t_id]:
                    technologies_dict[t_id].capex_anchors = dict(raw_tech_capex[t_id])
                    technologies_dict[t_id].capex_by_year = self._interpolate_dict(raw_tech_capex[t_id], years_list)
                if raw_tech_opex[t_id]:
                    technologies_dict[t_id].opex_anchors = dict(raw_tech_opex[t_id])
                    technologies_dict[t_id].opex_by_year = self._interpolate_dict(raw_tech_opex[t_id], years_list)
        
        specs_start = next((b['row'] for b in blocks_tech if b['type'] == 'START' and b['prefix'] == 'TECHNICAL SPECS'), None)
        specs_end = next((b['row'] for b in blocks_tech if b['type'] == 'END' and b['prefix'] == 'TECHNICAL SPECS'), None)
        if specs_start is not None and specs_end is not None:
            df_specs = self._extract_block_data(df_tech, specs_start, specs_end)
            for idx, row in df_specs.iterrows():
                if 'ID' in row.tolist() and 'RESSOURCE ID1' in row.tolist():
                    df_specs.columns = row.tolist()
                    df_specs = df_specs.iloc[idx+1:]
                    break
            for _, row in df_specs.iterrows():
                t_id = str(row.get('ID', '')).strip()
                res_id = str(row.get('RESSOURCE ID1', '')).strip()
                if t_id in technologies_dict and res_id and res_id != 'nan':
                    imp_type = str(row.get('TYPE (NEW/VARIATION)', '')).strip().lower()
                    val = row.get('VALUE', 0.0)
                    try: val = float(val)
                    except: val = 0.0
                    ref_resource = None
                    for col_candidate in ['RESSOURCE REF', 'RESOURCE REF', 'REF RESSOURCE', 'REF RESOURCE', 'RESSOURCE_REF', 'RESOURCE_REF', 'REF', 'REFERENCE RESSOURCE', 'REFERENCE RESOURCE']:
                        col_val = row.get(col_candidate, None)
                        if col_val is not None and str(col_val).strip() not in ('', 'nan'):
                            ref_resource = str(col_val).strip()
                            break
                    if ref_resource is None:
                        row_vals = [str(v).strip() for v in row.values if pd.notna(v)]
                        for v in row_vals:
                            if v in resources_dict and v != res_id:
                                ref_resource = v
                                break
                    technologies_dict[t_id].impacts[res_id] = {
                        'type': imp_type,
                        'value': val,
                        'reference': str(row.get('STATE: INITIAL/ACTUAL/AVOIDED', 'INITIAL')).strip().upper(),
                        'ref_resource': ref_resource
                    }
        
        tech_compatibilities = {}
        for i, row in df_tech.iterrows():
            row_vals = [str(x).strip().upper() for x in row if pd.notna(x)]
            if 'COMPATIBILITIES' in row_vals:
                if any(marker in row_vals for marker in ['START', 'END']): continue
                header_row = df_tech.iloc[i + 1]
                headers = []
                headers_start_col = -1
                for j, cell in enumerate(header_row):
                    val = str(cell).strip()
                    if val in technologies_dict:
                        headers.append(val)
                        if headers_start_col == -1: headers_start_col = j
                if headers_start_col != -1:
                    for k in range(i + 2, i + 2 + len(headers)):
                        if k >= len(df_tech): break
                        row_data = df_tech.iloc[k]
                        t_id_row = str(row_data.iloc[headers_start_col - 1]).strip()
                        if t_id_row in technologies_dict:
                            compat_dict = {}
                            for m, h_id in enumerate(headers):
                                cell_val = str(row_data.iloc[headers_start_col + m]).strip().upper()
                                if cell_val in ('X', 'FREE'):
                                    compat_dict[h_id] = cell_val
                            tech_compatibilities[t_id_row] = compat_dict
                break

        return technologies_dict, raw_tech_capex_links, raw_tech_opex_links, tech_compatibilities

    def _parse_entities(self, entities, entities_info, resources_dict, technologies_dict, scenario_id, active_sc_name, sc_name_map):
        import pandas as pd
        import numpy as np
        from .model import EntityState, Process
        import warnings
        entities_dict = {}
        for entity_id in entities:
            entity_meta = entities_info.get(entity_id, {})
            sheet_to_parse = str(entity_meta.get('sheet', '')).strip()
            if not sheet_to_parse:
                warnings.warn(f"Entity '{entity_id}' has no SHEET configured in CLUSTER block and will be skipped")
                continue
            if sheet_to_parse not in self.xl.sheet_names:
                warnings.warn(f"Entity '{entity_id}' references missing sheet '{sheet_to_parse}' and will be skipped")
                continue

            df_ent = self.xl.parse(sheet_to_parse, header=None)
            blocks_ent = self._find_blocks(df_ent)
            
            tot_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and b['prefix'] == 'TOTAL'), None)
            tot_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and b['prefix'] == 'TOTAL'), None)
            
            production_level = entity_meta.get('production', 0.0)
            init_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and b['prefix'] == 'INIT'), None)
            init_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and b['prefix'] == 'INIT'), None)
            annual_operating_hours = 8760.0
            sv_act_mode = "PI"
            if init_start is not None and init_end is not None:
                df_init = df_ent.iloc[init_start+1:init_end]
                tipe_op = 365.0
                hours_per_day = 24.0
                for idx, row in df_init.iterrows():
                    row_vals = [str(x).strip().upper() for x in row.values if pd.notna(x)]
                    if 'TIPE_OP' in row_vals:
                        idx_tipe = row_vals.index('TIPE_OP')
                        if len(row_vals) > idx_tipe + 1:
                            try: tipe_op = float(row_vals[idx_tipe+1])
                            except: pass
                        if len(row_vals) > idx_tipe + 2:
                            try: hours_per_day = float(row_vals[idx_tipe+2])
                            except: pass
                    if 'SV ACT' in row_vals:
                        idx_sv = row_vals.index('SV ACT')
                        if len(row_vals) > idx_sv + 1:
                            sv_act_mode = str(row_vals[idx_sv+1]).strip().upper()
                            if sv_act_mode not in ["PI", "NORM"]:
                                sv_act_mode = "PI"
                annual_operating_hours = tipe_op * hours_per_day
            
            annual_production = production_level * (annual_operating_hours / 24.0)
            
            process_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and 'PROCESS' in b['prefix']), None)
            process_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and 'PROCESS' in b['prefix']), None)
            processes_dict = {}
            if process_start is not None and process_end is not None:
                df_proc = df_ent.iloc[process_start+1:process_end]
                for _, row in df_proc.iterrows():
                    row_vals = [str(x).strip().upper() for x in row.values if pd.notna(x)]
                    p_id = ''
                    for v in row_vals:
                        if v.startswith('R') or v == 'R_OTHER':
                            p_id = v
                            break
                    if p_id:
                        p_name = p_id
                        raw_row_vals = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                        for v in raw_row_vals:
                            if v.upper() != p_id and not v.replace('.','').replace(',','').replace('%','').isnumeric() and v.upper() not in ['DES_PROCESS', 'NB_UNITS', '**']:
                                p_name = v
                                break
                        if p_id not in processes_dict:
                            p = Process(id=p_id, name=p_name)
                            processes_dict[p_id] = p
                        else:
                            p = processes_dict[p_id]
                        row_list = list(row.values)
                        if 'NB_UNITS' in row_vals:
                            try:
                                idx_id = raw_row_vals.index(p_id)
                                if idx_id + 1 < len(raw_row_vals):
                                    p.nb_units = int(float(raw_row_vals[idx_id + 1]))
                            except Exception:
                                pass
                        else:
                            for i, cell in enumerate(row_list):
                                cell_str = str(cell).strip()
                                if cell_str in resources_dict:
                                    try: 
                                        v = float(row_list[i+1])
                                        v = v if not np.isnan(v) else 0.0
                                        if self._is_primary_emission_resource(resources_dict[cell_str]):
                                            p.emission_shares[cell_str] = v
                                        else:
                                            p.consumption_shares[cell_str] = v
                                    except Exception:
                                        pass
            
            ca_percent = 0.0
            tech_trans_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and b['prefix'] == 'TECHNOLOGICAL TRANSITION'), None)
            tech_trans_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and b['prefix'] == 'TECHNOLOGICAL TRANSITION'), None)
            if tech_trans_start is not None and tech_trans_end is not None:
                df_tech_trans = df_ent.iloc[tech_trans_start+1:tech_trans_end]
                for idx, row in df_tech_trans.iterrows():
                    row_vals = [str(x).strip().upper() for x in row.values if pd.notna(x)]
                    if 'BUDGET' in row_vals and 'CA' in row_vals:
                        row_vals_raw = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                        has_budget = 'BUDGET' in row_vals
                        if has_budget and scenario_id:
                            sc_match = any(v.upper() in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"] for v in row_vals_raw)
                            if not sc_match: continue
                            else:
                                try:
                                    ca_idx = row_vals.index('CA')
                                    if len(row_vals) > ca_idx + 1:
                                        raw_ca = str(row_vals[ca_idx + 1]).strip()
                                        has_pct_sign = '%' in raw_ca
                                        ca_val = float(raw_ca.replace('%', ''))
                                        if has_pct_sign: ca_val /= 100.0
                                        elif ca_val > 1.0: ca_val /= 100.0
                                        ca_percent = ca_val
                                except Exception as e: pass
                        elif has_budget and not scenario_id:
                            try:
                                ca_idx = row_vals.index('CA')
                                if len(row_vals) > ca_idx + 1:
                                    raw_ca = str(row_vals[ca_idx + 1]).strip()
                                    has_pct_sign = '%' in raw_ca
                                    ca_val = float(raw_ca.replace('%', ''))
                                    if has_pct_sign: ca_val /= 100.0
                                    elif ca_val > 1.0: ca_val /= 100.0
                                    ca_percent = ca_val
                            except: pass
                    p_id_candidates = [v for v in row_vals if v in processes_dict]
                    if p_id_candidates:
                        p_id = p_id_candidates[0]
                        p_techs = [v for v in row_vals if v in technologies_dict]
                        existing_techs = set(processes_dict[p_id].valid_technologies)
                        new_techs = set(p_techs)
                        processes_dict[p_id].valid_technologies = list(existing_techs.union(new_techs))

            sold_resources = []
            purchases_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and b['prefix'] == 'PURCHASES'), None)
            purchases_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and b['prefix'] == 'PURCHASES'), None)
            if purchases_start is not None and purchases_end is not None:
                df_purchases = df_ent.iloc[purchases_start+1:purchases_end]
                for idx, row in df_purchases.iterrows():
                    row_vals = [str(x).strip().upper() for x in row.values if pd.notna(x)]
                    if 'SELL' in row_vals:
                        res_id = row_vals[row_vals.index('SELL') - 1]
                        sold_resources.append(res_id)
            
            base_cons = {}
            base_emis = 0.0
            
            ref_baselines = {}
            ref_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and 'REF' in b['prefix'].upper()), None)
            ref_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and 'REF' in b['prefix'].upper()), None)
            if ref_start is not None and ref_end is not None:
                for idx in range(ref_start + 1, ref_end):
                    row = df_ent.loc[idx]
                    res_col_idx = -1
                    val_col_idx = -1
                    year_col_idx = -1
                    for i, x in enumerate(row.values):
                        if pd.notna(x):
                            val = str(x).strip().upper()
                            if val in ('RESSOURCE ID', 'RESOURCE ID'): res_col_idx = i
                            elif val in ('VALOR', 'VALUE'): val_col_idx = i
                            elif val in ('YEAR', 'ANNEE', 'ANNÉE'): year_col_idx = i
                    if res_col_idx != -1 and val_col_idx != -1:
                        for j in range(idx + 1, ref_end):
                            drow = df_ent.loc[j]
                            row_vals_raw = [str(x).strip().upper() for x in drow.values if pd.notna(x)]
                            if scenario_id:
                                other_scs = [s for s in sc_name_map.keys() if s.upper() in row_vals_raw and s.upper() not in [scenario_id.upper(), str(active_sc_name).upper(), "ALL", "DEFAULT"]]
                                if other_scs: continue
                            r_id = str(drow.iloc[res_col_idx]).strip()
                            if pd.notna(drow.iloc[val_col_idx]) and r_id and r_id != 'nan':
                                yr = 2025
                                if year_col_idx != -1 and pd.notna(drow.iloc[year_col_idx]):
                                    try: yr = int(float(drow.iloc[year_col_idx]))
                                    except: pass
                                if r_id not in ref_baselines: ref_baselines[r_id] = {}
                                try: ref_baselines[r_id][yr] = float(drow.iloc[val_col_idx])
                                except Exception: pass
                        break
            
            if tot_start is not None and tot_end is not None:
                df_tot = df_ent.iloc[tot_start+1:tot_end]
                for idx, row in df_tot.iterrows():
                    vals = [x for x in row.values if pd.notna(x) and str(x).strip() != '']
                    if len(vals) >= 3:
                        r_id = str(vals[1]).strip()
                        val_str = vals[2]
                        raw_unit = str(vals[3]).strip().upper() if len(vals) > 3 else ''
                        try: val = float(val_str)
                        except: continue
                        multiplier = 1.0
                        if raw_unit == 'KGCO2': multiplier = 1 / 1000.0
                        elif raw_unit == 'KWH': multiplier = 1 / 1000.0
                        elif raw_unit == 'GJ': multiplier = 1 / 3.6
                        total_val = val * annual_production * multiplier
                        if r_id in resources_dict and self._is_primary_emission_resource(resources_dict[r_id]):
                            base_emis += total_val
                        elif r_id in resources_dict:
                            base_cons[r_id] = total_val
                            
            entities_dict[entity_id] = EntityState(
                id=entity_id, name=entity_meta.get('name', str(entity_id)), base_consumptions=base_cons, base_emissions=base_emis,
                production_level=annual_production, annual_operating_hours=annual_operating_hours,
                sv_act_mode=sv_act_mode, processes=processes_dict, ref_baselines=ref_baselines,
                ca_percentage_limit=ca_percent, sold_resources=sold_resources
            )
        return entities_dict

    def _parse_time_series(self, scenario_id, resources_dict, active_sc_name, sc_name_map, years_list):
        from .model import TimeSeriesData
        import pandas as pd
        time_series = TimeSeriesData()

        if 'RESSOURCES_PRICE' in self.xl.sheet_names:
            df_prices = self.xl.parse('RESSOURCES_PRICE', header=None)
            raw_prices = {}
            in_target_block = (scenario_id is None)
            found_any_scenario_block = False
            for _, row in df_prices.iterrows():
                row_vals_upper = [str(x).strip().upper() for x in row.values if pd.notna(x) and str(x).strip()]
                if 'START' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                    found_any_scenario_block = True
                    in_target_block = False
                    continue
                if 'END' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                    in_target_block = False
                    continue
                if 'SC-DES' in row_vals_upper and found_any_scenario_block:
                    raw = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                    raw_upper = [s.upper() for s in raw]
                    try:
                        sc_idx = raw_upper.index('SC-DES')
                        remaining = raw_upper[sc_idx+1:]
                        in_target_block = (scenario_id is None) or any(s in [scenario_id.upper(), active_sc_name] for s in remaining) or ("ALL" in remaining)
                    except: in_target_block = False
                    continue
                if not in_target_block: continue
                year = None
                start_idx = -1
                for idx, cell in enumerate(row.values):
                    try:
                        y = int(cell)
                        if y >= 2000 and y <= 2100:
                            year = y
                            start_idx = idx
                            break
                    except: pass
                if year is not None and start_idx != -1:
                    for i in range(start_idx + 1, len(row.values) - 1, 3):
                        r_id = str(row.values[i]).strip()
                        if r_id and r_id != 'nan' and pd.notna(row.values[i+1]):
                            try: price_val = float(row.values[i+1])
                            except: price_val = row.values[i+1]
                            if r_id not in raw_prices: raw_prices[r_id] = {}
                            raw_prices[r_id][year] = price_val
            for r_id, p_dict in raw_prices.items():
                if r_id not in resources_dict: raise ValueError(f"Resource '{r_id}' appears in RESSOURCES_PRICE but is missing from DATA block with a NAME")
                if p_dict:
                    time_series.resource_prices_anchors[r_id] = dict(p_dict)
                    time_series.resource_prices[r_id] = self._interpolate_dict(p_dict, years_list)

        if 'CARBON QUOTAS' in self.xl.sheet_names:
            df_cq = self.xl.parse('CARBON QUOTAS', header=None)
            raw_prices = {}
            raw_free_pi = {}
            raw_free_norm = {}
            raw_penalties = {}
            in_target_block = (scenario_id is None)
            found_any_scenario_block = False
            for _, row in df_cq.iterrows():
                row_vals_upper = [str(x).strip().upper() for x in row.values if pd.notna(x) and str(x).strip()]
                if 'START' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                    found_any_scenario_block = True
                    in_target_block = False
                    continue
                if 'END' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                    in_target_block = False
                    continue
                if 'SC-DES' in row_vals_upper and found_any_scenario_block:
                    raw = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                    raw_upper = [s.upper() for s in raw]
                    try:
                        sc_idx = raw_upper.index('SC-DES')
                        remaining = raw_upper[sc_idx+1:]
                        in_target_block = (scenario_id is None) or any(s in [scenario_id.upper(), active_sc_name] for s in remaining) or ("ALL" in remaining)
                    except: in_target_block = False
                    continue
                if not in_target_block: continue
                year = None
                start_idx = -1
                for idx, cell in enumerate(row.values):
                    try:
                        y = int(cell)
                        if y >= 2000 and y <= 2100:
                            year = y
                            start_idx = idx
                            break
                    except: pass
                if year is not None and start_idx != -1:
                    if len(row.values) > start_idx + 1 and pd.notna(row.values[start_idx+1]):
                        try: raw_prices[year] = float(row.values[start_idx+1])
                        except: raw_prices[year] = row.values[start_idx+1]
                    if len(row.values) > start_idx + 2 and pd.notna(row.values[start_idx+2]):
                        try: raw_penalties[year] = float(row.values[start_idx+2])
                        except: pass
                    if len(row.values) > start_idx + 3 and pd.notna(row.values[start_idx+3]):
                        try: raw_free_pi[year] = float(row.values[start_idx+3])
                        except: raw_free_pi[year] = row.values[start_idx+3]
                    if len(row.values) > start_idx + 4 and pd.notna(row.values[start_idx+4]):
                        try: raw_free_norm[year] = float(row.values[start_idx+4])
                        except: raw_free_norm[year] = row.values[start_idx+4]
            time_series.carbon_prices_anchors = dict(raw_prices)
            time_series.carbon_prices = self._interpolate_dict(raw_prices, years_list)
            time_series.carbon_quotas_pi = self._interpolate_dict(raw_free_pi, years_list)
            time_series.carbon_quotas_norm = self._interpolate_dict(raw_free_norm, years_list)
            interp_factors = self._interpolate_dict(raw_penalties, years_list)
            adjusted_factors = {}
            last_eff_price = -1.0
            for y in sorted(time_series.carbon_prices.keys()):
                p = time_series.carbon_prices[y]
                f = interp_factors.get(y, 0.0)
                eff_price = p * (1.0 + f)
                if eff_price < last_eff_price:
                    if p > 0: f = (last_eff_price / p) - 1.0
                    else: f = 0.0
                    eff_price = last_eff_price
                adjusted_factors[y] = f
                last_eff_price = eff_price
            time_series.carbon_penalties = adjusted_factors

        if 'OTHER EMISSIONS' in self.xl.sheet_names:
            df_oe = self.xl.parse('OTHER EMISSIONS', header=None)
            raw_ems = {}
            in_target_block_oe = (scenario_id is None)
            found_any_oe_block = False
            for _, row in df_oe.iterrows():
                row_vals_upper = [str(x).strip().upper() for x in row.values if pd.notna(x) and str(x).strip()]
                if 'START' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                    found_any_oe_block = True
                    in_target_block_oe = False
                    continue
                if 'END' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                    in_target_block_oe = False
                    continue
                if 'SC-DES' in row_vals_upper and found_any_oe_block:
                    raw = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                    raw_upper = [s.upper() for s in raw]
                    try:
                        sc_idx = raw_upper.index('SC-DES')
                        remaining = raw_upper[sc_idx+1:]
                        in_target_block_oe = (scenario_id is None) or any(s in [scenario_id.upper(), active_sc_name] for s in remaining) or ("ALL" in remaining)
                    except: in_target_block_oe = False
                    continue
                if not in_target_block_oe: continue
                year = None
                start_idx = -1
                for idx, cell in enumerate(row.values):
                    try:
                        y = int(cell)
                        if y >= 2000 and y <= 2100:
                            year = y
                            start_idx = idx
                            break
                    except: pass
                if year is not None and start_idx != -1:
                    for i in range(start_idx + 1, len(row.values) - 1, 3):
                        r_id = str(row.values[i]).strip()
                        if r_id and r_id != 'nan' and pd.notna(row.values[i+1]):
                            em_val = row.values[i+1]
                            if r_id not in raw_ems: raw_ems[r_id] = {}
                            raw_ems[r_id][year] = em_val
            for r_id, e_dict in raw_ems.items():
                time_series.other_emissions_factors[r_id] = self._interpolate_dict(e_dict, years_list)
        return time_series

    def _parse_public_aids(self):
        import pandas as pd
        from .model import GrantParams, CCfDParams
        grant_params = GrantParams()
        ccfd_params = CCfDParams()
        if 'PUBLIC AID' in self.xl.sheet_names:
            df_aid = self.xl.parse('PUBLIC AID', header=None)
            blocks_aid = self._find_blocks(df_aid)
            init_start = next((b['row'] for b in blocks_aid if b['type'] == 'START' and b['prefix'] == 'INIT'), None)
            init_end = next((b['row'] for b in blocks_aid if b['type'] == 'END' and b['prefix'] == 'INIT'), None)
            active = False
            if init_start is not None and init_end is not None:
                df_init = df_aid.iloc[init_start+1:init_end]
                for _, row in df_init.iterrows():
                    row_list = list(row.values)
                    vals = [str(x).strip().upper() for x in row_list if pd.notna(x)]
                    if 'ACTIVE' in vals:
                        active_idx = -1
                        for i, x in enumerate(row_list):
                            if str(x).strip().upper() == 'ACTIVE':
                                active_idx = i
                                break
                        if active_idx != -1 and len(row_list) > active_idx + 1:
                            if str(row_list[active_idx+1]).strip().upper() == 'YES':
                                active = True
            if active:
                inc_start = next((b['row'] for b in blocks_aid if b['type'] == 'START' and 'INCENTIVES' in b['prefix']), None)
                inc_end = next((b['row'] for b in blocks_aid if b['type'] == 'END' and 'INCENTIVES' in b['prefix']), None)
                if inc_start is not None and inc_end is not None:
                    df_inc = df_aid.iloc[inc_start+1:inc_end]
                    for idx, row in df_inc.iterrows():
                        row_list = list(row.values)
                        vals = [str(x).strip().upper() for x in row_list if pd.notna(x)]
                        if 'GRANT' in vals:
                            grant_idx = -1
                            for i, x in enumerate(row_list):
                                if str(x).strip().upper() == 'GRANT':
                                    grant_idx = i
                                    break
                            if grant_idx != -1:
                                if len(row_list) > grant_idx + 1 and pd.notna(row_list[grant_idx + 1]):
                                    try: grant_params.rate = float(str(row_list[grant_idx + 1]).replace('%', '')) / 100.0 if '%' in str(row_list[grant_idx + 1]) else float(row_list[grant_idx + 1])
                                    except: pass
                                if len(row_list) > grant_idx + 2 and pd.notna(row_list[grant_idx + 2]):
                                    try: grant_params.cap = float(row_list[grant_idx + 2])
                                    except: pass
                                if len(row_list) > grant_idx + 3 and pd.notna(row_list[grant_idx + 3]):
                                    try: grant_params.renew_time = float(row_list[grant_idx + 3])
                                    except: pass
                                grant_params.active = True
                        if 'SUBS_NO' in vals:
                            subs_no_idx = -1
                            for i, x in enumerate(row_list):
                                if str(x).strip().upper() == 'SUBS_NO':
                                    subs_no_idx = i
                                    break
                            if subs_no_idx != -1:
                                for i in range(subs_no_idx + 1, len(row_list)):
                                    val_tech = str(row_list[i]).strip().upper()
                                    if pd.notna(row_list[i]) and val_tech and val_tech != 'NAN':
                                        grant_params.excluded_technologies.append(val_tech)
                        if 'CCFD' in vals:
                            ccfd_idx = -1
                            for i, x in enumerate(row_list):
                                if str(x).strip().upper() == 'CCFD':
                                    ccfd_idx = i
                                    break
                            if ccfd_idx != -1:
                                if len(row_list) > ccfd_idx + 1 and pd.notna(row_list[ccfd_idx + 1]):
                                    try: ccfd_params.duration = int(float(row_list[ccfd_idx + 1]))
                                    except: pass
                                if len(row_list) > ccfd_idx + 2 and pd.notna(row_list[ccfd_idx + 2]):
                                    try: ccfd_params.contract_type = int(float(row_list[ccfd_idx + 2]))
                                    except: pass
                                if len(row_list) > ccfd_idx + 3 and pd.notna(row_list[ccfd_idx + 3]):
                                    try: ccfd_params.eua_price_pct = float(str(row_list[ccfd_idx + 3]).replace('%', '')) / 100.0 if '%' in str(row_list[ccfd_idx + 3]) else float(row_list[ccfd_idx + 3])
                                    except: pass
                                if len(row_list) > ccfd_idx + 4 and pd.notna(row_list[ccfd_idx + 4]):
                                    try: ccfd_params.nb_contracts = int(float(row_list[ccfd_idx + 4]))
                                    except: pass
                                ccfd_params.active = True
        return grant_params, ccfd_params

    def _parse_bank_loans(self, duration):
        import pandas as pd
        from .model import BankLoan
        bank_loans = []
        if 'BANK' in self.xl.sheet_names:
            df_bank = self.xl.parse('BANK', header=None)
            blocks_bank = self._find_blocks(df_bank)
            prod_start = next((b['row'] for b in blocks_bank if b['type'] == 'START' and b['prefix'] == 'PRODUCTS'), None)
            prod_end = next((b['row'] for b in blocks_bank if b['type'] == 'END' and b['prefix'] == 'PRODUCTS'), None)
            if prod_start is not None and prod_end is not None:
                df_prod = self._extract_block_data(df_bank, prod_start, prod_end)
                for idx, row in df_prod.iterrows():
                    row_list = [str(x).upper().strip() for x in row.tolist()]
                    if 'RATE (%)' in row_list or 'LOAN PERIOD (YEARS)' in row_list:
                        df_prod.columns = row_list
                        df_prod = df_prod.iloc[idx+1:]
                        break
                for _, row in df_prod.iterrows():
                    try:
                        rate_raw = str(row.get('RATE (%)', '0')).replace('%', '').strip()
                        rate = float(rate_raw) / 100.0 if rate_raw else 0.0
                        duration_val_raw = str(row.get('LOAN PERIOD (YEARS)', '1')).strip().upper()
                        if duration_val_raw == 'ALL':
                            for d in range(1, duration + 1):
                                bank_loans.append(BankLoan(rate=rate, duration=d))
                        else:
                            loan_duration = int(float(duration_val_raw))
                            if loan_duration < 1: loan_duration = 1
                            bank_loans.append(BankLoan(rate=rate, duration=loan_duration))
                    except: pass
        return bank_loans

    def _parse_dac_and_credits(self, scenario_id, active_sc_name, sc_name_map, years_list):
        import pandas as pd
        from .model import DACParams, CreditParams
        dac_params = DACParams()
        credit_params = CreditParams()
        if 'NEW TECH_INDIRECT' in self.xl.sheet_names:
            df_ind = self.xl.parse('NEW TECH_INDIRECT', header=None)
            in_dac = False
            in_credit = False
            raw_dac_capex = {}
            raw_dac_opex_pct = {}
            raw_dac_elec = {}
            raw_credit_cost = {}
            for _, row in df_ind.iterrows():
                row_vals = [str(x).strip().upper() for x in row.values if pd.notna(x)]
                if not row_vals: continue
                if ('DAC' in row_vals) or ('DIRECT AIR CAPTURE' in row_vals):
                    in_dac = True
                    dac_params.active = True
                if ('CREDIT' in row_vals or 'CARBON CREDIT' in row_vals) and ('START' in row_vals):
                    in_credit = True
                    credit_params.active = True
                if 'DAC' in row_vals and 'END' in row_vals: in_dac = False
                if 'CREDIT' in row_vals and 'END' in row_vals: in_credit = False
                if in_dac:
                    if 'ACT' in row_vals:
                        row_list = list(row.values)
                        try:
                            idx = next((i for i, x in enumerate(row_list) if str(x).strip().upper() == 'ACT'), -1)
                            if idx != -1:
                                row_sc = str(row_list[idx-1]).strip().upper() if idx > 0 and pd.notna(row_list[idx-1]) else ''
                                if scenario_id and row_sc and row_sc not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]: continue
                                if len(row_list) > idx + 1 and pd.notna(row_list[idx+1]): dac_params.active = (str(row_list[idx+1]).strip().upper() == 'YES')
                                if len(row_list) > idx + 2 and pd.notna(row_list[idx+2]):
                                    val = str(row_list[idx+2]).strip()
                                    if val and val.replace('.', '', 1).isnumeric(): dac_params.start_year = int(float(val))
                                if len(row_list) > idx + 3 and pd.notna(row_list[idx+3]):
                                    val = str(row_list[idx+3]).strip()
                                    if val and val.replace('.', '', 1).isnumeric(): dac_params.end_year = int(float(val))
                        except: pass
                    if 'CARAC' in row_vals:
                        vals = [x for x in row.values if pd.notna(x)]
                        if len(vals) >= 7 and str(vals[0]).strip().upper() == 'CARAC':
                            carac_val1 = str(vals[1]).strip()
                            carac_scenario = None
                            carac_start_idx = 1
                            try: int(float(carac_val1))
                            except (ValueError, TypeError):
                                carac_scenario = carac_val1.upper()
                                carac_start_idx = 2
                            if scenario_id and carac_scenario and carac_scenario not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]: continue
                            carac_vals = vals[carac_start_idx:]
                            if len(carac_vals) < 6: continue
                            try:
                                year = int(carac_vals[0])
                                try: raw_dac_capex[year] = float(carac_vals[1])
                                except: raw_dac_capex[year] = str(carac_vals[1]).strip()
                                try: raw_dac_opex_pct[year] = float(carac_vals[4])
                                except: raw_dac_opex_pct[year] = str(carac_vals[4]).strip()
                                try: raw_dac_elec[year] = float(carac_vals[5])
                                except: raw_dac_elec[year] = str(carac_vals[5]).strip()
                            except Exception: pass
                if in_credit:
                    if 'ACT' in row_vals:
                        row_list = list(row.values)
                        try:
                            idx = next((i for i, x in enumerate(row_list) if str(x).strip().upper() == 'ACT'), -1)
                            if idx != -1:
                                row_sc = str(row_list[idx-1]).strip().upper() if idx > 0 and pd.notna(row_list[idx-1]) else ''
                                if scenario_id and row_sc and row_sc not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]: continue
                                if len(row_list) > idx + 1 and pd.notna(row_list[idx+1]): credit_params.active = (str(row_list[idx+1]).strip().upper() == 'YES')
                                if len(row_list) > idx + 2 and pd.notna(row_list[idx+2]):
                                    val = str(row_list[idx+2]).strip()
                                    if val and val.replace('.', '', 1).isnumeric(): credit_params.start_year = int(float(val))
                                if len(row_list) > idx + 3 and pd.notna(row_list[idx+3]):
                                    val = str(row_list[idx+3]).strip()
                                    if val and val.replace('.', '', 1).isnumeric(): credit_params.end_year = int(float(val))
                        except: pass
                    if 'CREDIT' in row_vals and 'START' not in row_vals and 'END' not in row_vals:
                        vals = [x for x in row.values if pd.notna(x)]
                        if len(vals) >= 5 and str(vals[0]).strip().upper() == 'CREDIT':
                            credit_sc = str(vals[2]).strip().upper() if len(vals) > 2 else ''
                            if scenario_id and credit_sc and credit_sc not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]: continue
                            try:
                                if len(vals) >= 6:
                                    year = int(float(vals[4]))
                                    cost = float(str(vals[5]).strip().replace('€', '').replace('$', '').replace(',', '.').strip())
                                    raw_credit_cost[year] = cost
                                elif len(vals) >= 5:
                                    year = int(float(vals[3]))
                                    cost = float(str(vals[4]).replace('€', '').replace(',', '.').strip())
                                    raw_credit_cost[year] = cost
                            except Exception: pass
                if 'TREHS' in row_vals:
                    vals = [x for x in row.values if pd.notna(x)]
                    if len(vals) >= 3 and str(vals[0]).strip().upper() == 'TREHS':
                        v_sc = None
                        v_year = 2025
                        v_pct = 100.0
                        potential_pcts = []
                        potential_years = []
                        potential_scs = []
                        for v in vals[1:]:
                            v_str = str(v).strip()
                            if isinstance(v, str) and '%' in v_str:
                                try: potential_pcts.append(float(v_str.replace('%', '')) / 100.0)
                                except: pass
                                continue
                            try:
                                fv = float(v_str)
                                if (fv >= 1900 and fv <= 2100) and v_year == 2025: potential_years.append(int(fv))
                                elif fv <= 100.0: potential_pcts.append(fv)
                                continue
                            except: pass
                            if v_str.upper() in [s.upper() for s in sc_name_map.keys()] or v_str.upper() in [s.upper() for s in sc_name_map.values()] or v_str.upper() in ["ALL", "DEFAULT"]:
                                potential_scs.append(v_str.upper())
                        if potential_years: v_year = potential_years[0]
                        if potential_pcts:
                            v_pct = potential_pcts[0]
                            if v_pct > 1.0: v_pct = v_pct / 100.0
                        if potential_scs: v_sc = potential_scs[0]
                        if scenario_id and v_sc and v_sc not in [scenario_id.upper(), str(active_sc_name).upper(), "ALL", "DEFAULT"]: continue
                        if in_dac:
                            dac_params.ref_year = v_year
                            dac_params.max_volume_pct = v_pct
                        elif in_credit:
                            credit_params.ref_year = v_year
                            credit_params.max_volume_pct = v_pct

            if dac_params.active:
                dac_params.capex_by_year = self._interpolate_dict(raw_dac_capex, years_list)
                if not dac_params.capex_by_year: dac_params.active = False
                else:
                    interp_opex_pct = self._interpolate_dict(raw_dac_opex_pct, years_list)
                    dac_params.opex_by_year = {y: dac_params.capex_by_year.get(y, 0.0) * (interp_opex_pct.get(y, 0.0) / 100.0) for y in dac_params.capex_by_year}
                    dac_params.elec_by_year = self._interpolate_dict(raw_dac_elec, years_list)
            if credit_params.active:
                credit_params.cost_by_year = self._interpolate_dict(raw_credit_cost, years_list)
                if not credit_params.cost_by_year: credit_params.active = False
        return dac_params, credit_params

    def parse(self, scenario_id: str = None) -> PathFinderData:
        sc_meta = []
        try: sc_meta = self._parse_scenarios()
        except: pass
        sc_name_map = {s['id'].upper(): s['name'].upper() for s in sc_meta}
        active_sc_name = sc_name_map.get(scenario_id.upper() if scenario_id else "", "")

        df_overview = self.xl.parse('OverView', header=None)
        blocks_overview = self._find_blocks(df_overview)
        
        reporting_toggles, start_year, duration, time_limit, mip_gap, relax_integrality, discount_rate, run_project = self._parse_overview_settings(df_overview, blocks_overview)
        years_list = list(range(start_year, start_year + duration + 1))
        
        entities_info, entities = self._parse_entities_cluster(df_overview, blocks_overview)
        resources_dict = self._parse_resources(df_overview, blocks_overview)
        unit_conversions = self._parse_unit_conversions(df_overview, blocks_overview)
        objectives_list = self._parse_objectives(df_overview, blocks_overview)

        params = Parameters(
            start_year=start_year, duration=duration, entities=entities, resources=list(resources_dict.keys()),
            time_limit=time_limit, mip_gap=mip_gap, relax_integrality=relax_integrality, discount_rate=discount_rate,
            run_project=run_project
        )

        df_tech = self.xl.parse('NEW TECH', header=None)
        technologies_dict, raw_tech_capex_links, raw_tech_opex_links, tech_compatibilities = self._parse_technologies(
            df_tech, scenario_id, active_sc_name, resources_dict, years_list
        )

        entities_dict = self._parse_entities(entities, entities_info, resources_dict, technologies_dict, scenario_id, active_sc_name, sc_name_map)
        
        time_series = self._parse_time_series(scenario_id, resources_dict, active_sc_name, sc_name_map, years_list)
        
        grant_params, ccfd_params = self._parse_public_aids()
        bank_loans = self._parse_bank_loans(duration)
        dac_params, credit_params = self._parse_dac_and_credits(scenario_id, active_sc_name, sc_name_map, years_list)

        for t_id in technologies_dict:
            if t_id in raw_tech_capex_links:
                link_id, base_cost = raw_tech_capex_links[t_id]
                linked_prices = time_series.carbon_prices if link_id == 'EUA' else time_series.resource_prices.get(link_id, {})
                if linked_prices:
                    baseline_p = linked_prices.get(params.start_year, 1.0)
                    if baseline_p == 0: baseline_p = 1.0
                    technologies_dict[t_id].capex_by_year = {y: base_cost * (linked_prices.get(y, baseline_p) / baseline_p) for y in years_list}
            if t_id in raw_tech_opex_links:
                link_id, base_cost = raw_tech_opex_links[t_id]
                linked_prices = time_series.carbon_prices if link_id == 'EUA' else time_series.resource_prices.get(link_id, {})
                if linked_prices:
                    baseline_p = linked_prices.get(params.start_year, 1.0)
                    if baseline_p == 0: baseline_p = 1.0
                    technologies_dict[t_id].opex_by_year = {y: base_cost * (linked_prices.get(y, baseline_p) / baseline_p) for y in years_list}

        return PathFinderData(
            parameters=params, resources=resources_dict, technologies=technologies_dict,
            time_series=time_series, entities=entities_dict, objectives=objectives_list,
            tech_compatibilities=tech_compatibilities, unit_conversions=unit_conversions,
            grant_params=grant_params, ccfd_params=ccfd_params, bank_loans=bank_loans,
            dac_params=dac_params, credit_params=credit_params, reporting_toggles=reporting_toggles
        )

    # ─────────────────────────────────────────────────────────────────────────
    # ANALYSE DE SENSIBILITÉ
    # ─────────────────────────────────────────────────────────────────────────

    def _parse_sensitivity_block(self, df_overview: pd.DataFrame) -> SensitivityParams:
        """
        Analyse le bloc SENSITIVITY START / SENSITIVITY END de la feuille OverView
        et retourne un objet SensitivityParams complet.

        Structure attendue (tags en colonne A) :
            VAR     → amplitudes de variation (%, ex: 5% 10% 25% 50% 100%)
            P/N     → direction (P, N ou ALL)
            SIM     → scénarios cibles (ex: BS)
            TIME    → temps limite (s) par simulation
            DATA?   → EUA YES/NO, RESSOURCES PRICE YES/NO, …
            INDI    → nom d'un indicateur à surveiller
        """
        variations: List[float] = []
        run: bool = False
        direction: str = "ALL"
        scenarios: List[str] = []
        time_limit: int = 10
        targets: Dict[str, bool] = {}
        indicators: List[str] = []

        in_block = False

        for _, row in df_overview.iterrows():
            # Valeurs brutes et normalisées
            raw_vals = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
            vals_upper = [v.upper() for v in raw_vals]

            # Détection des marqueurs START / END du bloc SENSITIVITY
            if 'SENSITIVITY' in vals_upper and 'START' in vals_upper:
                in_block = True
                continue
            if 'SENSITIVITY' in vals_upper and 'END' in vals_upper:
                break

            if not in_block:
                continue

            # Ignorer les lignes de commentaire ou vides
            if not raw_vals or raw_vals[0].startswith('**'):
                continue

            tag = vals_upper[0]

            # ── Commande de lancement (RUN YES/NO) ──────────────────────────
            if tag == 'RUN':
                for token in raw_vals[1:]:
                    val = token.strip().upper()
                    if val in ('YES', 'NO'):
                        run = (val == 'YES')
                        break

            # ── Amplitudes de variation ────────────────────────────────────
            elif tag == 'VAR':
                for token in raw_vals[1:]:
                    token_clean = token.replace('%', '').strip()
                    try:
                        val = float(token_clean)
                        # Convertir les pourcentages entiers en décimaux
                        if val > 1.0:
                            val /= 100.0
                        if val > 0:
                            variations.append(round(val, 6))
                    except ValueError:
                        pass

            # ── Direction P / N / ALL ──────────────────────────────────────
            elif tag == 'P/N':
                for token in raw_vals[1:]:
                    t_up = token.strip().upper()
                    if t_up in ('P', 'N', 'ALL'):
                        direction = t_up
                        break

            # ── Scénarios à simuler ────────────────────────────────────────
            elif tag == 'SIM':
                placeholder_skip = ('SC1', 'SC2', 'SC3', 'SC4', 'SC...', 'SC…')
                for token in raw_vals[1:]:
                    if token.upper() not in placeholder_skip:
                        scenarios.append(token.upper())

            # ── Temps limite par simulation ────────────────────────────────
            elif tag == 'TIME' or 'TIME' in tag:
                for token in raw_vals[1:]:
                    try:
                        time_limit = int(float(token))
                        break
                    except ValueError:
                        pass

            # ── Données à perturber (DATA?) ────────────────────────────────
            elif tag == 'DATA?':
                # Format attendu : DATA? | <NOM_PARAMETRE> | YES/NO
                if len(raw_vals) >= 3:
                    param_name = raw_vals[1].strip()
                    yn_token = raw_vals[2].strip().upper()
                    targets[param_name] = (yn_token == 'YES')
                elif len(raw_vals) == 2:
                    # Valeur YES/NO absente, supposée FALSE
                    targets[raw_vals[1].strip()] = False

            # ── Indicateurs KPI ────────────────────────────────────────────
            elif tag == 'INDI':
                # Format attendu : INDI | <NOM_INDICATEUR>
                if len(raw_vals) >= 2:
                    indi_name = raw_vals[1].strip()
                    if indi_name and indi_name.upper() not in ('NOM', 'NAME', 'INDICATOR'):
                        indicators.append(indi_name)

        return SensitivityParams(
            run=run,
            variations=variations,
            direction=direction,
            scenarios=scenarios,
            time_limit=time_limit,
            targets=targets,
            indicators=indicators,
        )

    def _interpolate_dict(self, data_dict: Dict[int, Any], years_list: List[int]) -> Dict[int, float]:
        """Utility to interpolate a dictionary of {year: value} over a full years_list."""
        if not data_dict: return {}
        s = pd.Series(data_dict).sort_index()
        s = s.reindex(years_list)
        return self._interpolate_linear(s).to_dict()

    def parse_sensitivity(self) -> SensitivityParams:
        """
        Point d'entrée public pour lire les paramètres d'analyse de sensibilité
        depuis le fichier Excel sans déclencher le parseur complet des scénarios.

        Utilisation :
            parser = PathFinderParser('PathFinder input.xlsx')
            sens_params = parser.parse_sensitivity()
        """
        df_overview = self.xl.parse('OverView', header=None)
        params = self._parse_sensitivity_block(df_overview)

        if self.verbose:
            pass # Sensitivity parameter logs removed for cleaner terminal

        return params


if __name__ == '__main__':
    parser = PathFinderParser('data/raw/excel/PathFinder input.xlsx')
    data = parser.parse()
