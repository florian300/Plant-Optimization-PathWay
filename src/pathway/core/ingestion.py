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

    def _parse_numeric(self, val, default=0.0) -> float:
        """Nettoie et convertit une valeur en float, gérant les devises, les séparateurs et les pourcentages."""
        if pd.isna(val) or val == '' or str(val).strip().lower() == 'nan':
            return default
        
        val_str = str(val).strip()
        # Nettoyage des symboles monétaires et d'échelle
        for char in ['€', '$', 'M', ' ']:
            val_str = val_str.replace(char, '')
        
        # Gestion du séparateur décimal (remplacement de la virgule par un point)
        val_str = val_str.replace(',', '.')
        
        try:
            has_pct = '%' in val_str
            num_val = float(val_str.replace('%', ''))
            if has_pct:
                num_val /= 100.0
            return num_val
        except (ValueError, TypeError):
            return default

    def _parse_bool(self, raw_val) -> bool:
        """Convertit une valeur brute en booléen (gère YES/NO, TRUE/FALSE, 1/0)."""
        if pd.isna(raw_val):
            return False
        v = str(raw_val).strip().upper()
        if v in ['YES', 'TRUE', '1', 'Y']:
            return True
        if v in ['NO', 'FALSE', '0', 'N', '', 'NAN']:
            return False
        try:
            return float(v) != 0.0
        except (ValueError, TypeError):
            return False

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
                        is_yes = self._parse_bool(next((v for v in row_vals if v in ('YES', 'NO')), 'NO'))
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
                            num = self._parse_numeric(val, None)
                            if num is not None:
                                reporting_toggles.investment_cap = num
                                break
        
        start_year, duration, time_limit, mip_gap, relax_integrality, discount_rate, run_project = 2025, 25, 60.0, 0.90, False, 0.0, True
        for i, row in df_overview.iterrows():
            row_vals = [str(x).strip().upper() if pd.notna(x) else "" for x in row]
            if "YEAR START" in row_vals:
                idx = row_vals.index("YEAR START")
                if len(row) > idx + 1:
                    start_year = int(self._parse_numeric(row.iloc[idx + 1], 2025))
            if "SIMULATION TIME (IN YEAR)" in row_vals:
                idx = row_vals.index("SIMULATION TIME (IN YEAR)")
                if len(row) > idx + 1:
                    duration = int(self._parse_numeric(row.iloc[idx + 1], 25))
            if "DURURATION SIMULATION (S)" in row_vals or "DURATION SIMULATION (S)" in row_vals:
                keyword = "DURURATION SIMULATION (S)" if "DURURATION SIMULATION (S)" in row_vals else "DURATION SIMULATION (S)"
                idx = row_vals.index(keyword)
                if len(row) > idx + 1:
                    time_limit = self._parse_numeric(row.iloc[idx + 1], 60.0)
            if "RELAX INTEGRAL" in row_vals:
                idx = row_vals.index("RELAX INTEGRAL")
                if len(row) > idx + 1:
                    relax_integrality = self._parse_bool(row.iloc[idx + 1])
            if "ERROR SIMULATION (%)" in row_vals:
                idx = row_vals.index("ERROR SIMULATION (%)")
                if len(row) > idx + 1:
                    mip_gap = self._parse_numeric(row.iloc[idx + 1], 0.90)
            if "DISCOUNT RATE (%)" in row_vals or "DISCOUNT RATE" in row_vals:
                keyword = "DISCOUNT RATE (%)" if "DISCOUNT RATE (%)" in row_vals else "DISCOUNT RATE"
                idx = row_vals.index(keyword)
                if len(row) > idx + 1:
                    discount_rate = self._parse_numeric(row.iloc[idx + 1], 0.0)
            if "RUN PROJECT ?" in row_vals or "RUN PROJECT ? (YES/NO)" in row_vals:
                keyword = "RUN PROJECT ?" if "RUN PROJECT ?" in row_vals else "RUN PROJECT ? (YES/NO)"
                idx = row_vals.index(keyword)
                if len(row) > idx + 1:
                    run_project = self._parse_bool(row.iloc[idx + 1])
        return reporting_toggles, start_year, duration, time_limit, mip_gap, relax_integrality, discount_rate, run_project

    def _parse_entities_cluster(self, df_overview, blocks_overview):
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
                        prod = self._parse_numeric(row.get(prod_col, 0.0), 0.0)
                        sheet_name = str(row.get(sheet_col, '')).strip() if sheet_col is not None else ''
                        entity_name = str(row.get(name_col, '')).strip() if name_col is not None else str(e_id)
                        if not entity_name or entity_name.lower() == 'nan': entity_name = e_id
                        entities_info[e_id] = {'production': prod, 'sheet': sheet_name, 'name': entity_name}
                entities = list(entities_info.keys())
        return entities_info, entities

    def _parse_resources(self, df_overview, blocks_overview):
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
                        
                        # New column for indirect carbon tax
                        tax_indir = self._parse_bool(row.get('CARBON TAX ON INDIRECT EMISSIONS ? (YES/NO)', 'NO'))
                        
                        resources_dict[res_id] = Resource(
                            id=res_id, 
                            type=str(row['TYPE']), 
                            unit=str(row['UNIT']), 
                            name=res_name, 
                            category=category, 
                            resource_type=resource_type,
                            tax_indirect_emissions=tax_indir
                        )
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
                    factor = self._parse_numeric(factor_raw, None)
                    if factor is None: raise ValueError(f"Invalid FACTOR in UNIT CONVERSIONS at row {row_idx}: {factor_raw}")
                    if factor == 0.0: raise ValueError(f"UNIT CONVERSIONS factor cannot be zero for {unit_in} -> {unit_out}")
                    unit_conversions[(unit_in, unit_out)] = factor
                    unit_conversions[(unit_out, unit_in)] = 1.0 / factor
        return unit_conversions

    def _parse_objectives(self, df_overview, blocks_overview):
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
                        t_year = int(self._parse_numeric(row.get('TARGET_YEAR', 0), 0))
                        c_val = self._parse_numeric(row.get('CAP_VALUE', 0.0), 0.0)
                        c_year = None
                        try: 
                            if pd.notna(row.get('COMP_YEAR')):
                                c_year = int(self._parse_numeric(row.get('COMP_YEAR', 0), 0))
                        except (ValueError, TypeError): pass
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
            for _, row in df_tecs.iterrows():
                t_id = str(row.get('ID', '')).strip()
                if t_id and t_id != 'nan':
                    imp_time_raw = row.get('IMPLEMANTATION TIME (YEAR)') or row.get('IMPLEMENTATION TIME (YEAR)')
                    imp_time = int(self._parse_numeric(imp_time_raw, 1))
                    t_name = str(row.get('NAME', '')).strip()
                    if not t_name or t_name.lower() == 'nan': raise ValueError(f"Technology '{t_id}' is missing NAME in TECS block")
                    is_continuous_improvement = False
                    if is_ci_col is not None: is_continuous_improvement = self._parse_bool(row.get(is_ci_col, False))
                    if t_id.upper() == 'UP': is_continuous_improvement = True
                    tech_category = "Standard"
                    if tech_category_col is not None:
                        raw_cat = str(row.get(tech_category_col, '')).strip()
                        if raw_cat and raw_cat.lower() != 'nan': tech_category = raw_cat
                    t_id_upper = t_id.upper()
                    if tech_category == 'Standard' and ('CCS' in t_id_upper or 'CCU' in t_id_upper):
                        tech_category = 'Carbon Capture'
                    technologies_dict[t_id] = Technology(id=t_id, name=t_name, implementation_time=imp_time, capex=0.0, opex=0.0, impacts={}, is_continuous_improvement=is_continuous_improvement, tech_category=tech_category)
        
        euro_start = next((b['row'] for b in blocks_tech if b['type'] == 'START' and 'SPECS' in b['prefix'] and 'TECHNICAL' not in b['prefix']), None)
        euro_end = next((b['row'] for b in blocks_tech if b['type'] == 'END' and 'SPECS' in b['prefix'] and 'TECHNICAL' not in b['prefix']), None)
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
                    cost_val = self._parse_numeric(row.get('COST', 0), 0.0)
                    y_str = str(row.get('YEAR', 'ALL')).strip().upper()
                    try: year_val = int(float(y_str))
                    except (ValueError, TypeError): year_val = y_str
                    per_unit_str = str(row.get('PER UNIT ?', 'NO')).strip().upper()
                    is_per_unit = self._parse_bool(per_unit_str)
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
                    val = self._parse_numeric(row.get('VALUE', 0.0), 0.0)
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
        compat_start = next((b['row'] for b in blocks_tech if b['type'] == 'START' and 'COMPATIBILITIES' in b['prefix']), None)
        compat_end = next((b['row'] for b in blocks_tech if b['type'] == 'END' and 'COMPATIBILITIES' in b['prefix']), None)
        
        if compat_start is not None and compat_end is not None:
            df_compat = self._extract_block_data(df_tech, compat_start, compat_end)
            if not df_compat.empty:
                headers = [str(c).strip() for c in df_compat.columns if str(c).strip() in technologies_dict]
                if headers:
                    for _, row in df_compat.iterrows():
                        # The first column usually contains the row technology ID
                        t_id_row = str(row.iloc[0]).strip()
                        if t_id_row in technologies_dict:
                            compat_dict = {}
                            for h_id in headers:
                                cell_val = str(row.get(h_id, '')).strip().upper()
                                if cell_val in ('X', 'FREE'):
                                    compat_dict[h_id] = cell_val
                            if compat_dict:
                                tech_compatibilities[t_id_row] = compat_dict

        return technologies_dict, raw_tech_capex_links, raw_tech_opex_links, tech_compatibilities

    def _parse_entities(self, entities, entities_info, resources_dict, technologies_dict, scenario_id, active_sc_name, sc_name_map):
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
                            tipe_op = self._parse_numeric(row_vals[idx_tipe+1], 365.0)
                        if len(row_vals) > idx_tipe + 2:
                            hours_per_day = self._parse_numeric(row_vals[idx_tipe+2], 24.0)
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
                                    p.nb_units = int(self._parse_numeric(raw_row_vals[idx_id + 1], 0))
                            except (ValueError, TypeError):
                                pass
                        else:
                            for i, cell in enumerate(row_list):
                                cell_str = str(cell).strip()
                                if cell_str in resources_dict:
                                    v = self._parse_numeric(row_list[i+1], 0.0)
                                    if self._is_primary_emission_resource(resources_dict[cell_str]):
                                        p.emission_shares[cell_str] = v
                                    else:
                                        p.consumption_shares[cell_str] = v
            
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
                                        ca_percent = self._parse_numeric(row_vals[ca_idx + 1], 0.0)
                                except (ValueError, TypeError): pass
                        elif has_budget and not scenario_id:
                            try:
                                ca_idx = row_vals.index('CA')
                                if len(row_vals) > ca_idx + 1:
                                    ca_percent = self._parse_numeric(row_vals[ca_idx + 1], 0.0)
                            except (ValueError, TypeError): pass
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
                                    yr = int(self._parse_numeric(drow.iloc[year_col_idx], 2025))
                                if r_id not in ref_baselines: ref_baselines[r_id] = {}
                                ref_baselines[r_id][yr] = self._parse_numeric(drow.iloc[val_col_idx], 0.0)
                        break
            
            if tot_start is not None and tot_end is not None:
                df_tot = df_ent.iloc[tot_start+1:tot_end]
                for idx, row in df_tot.iterrows():
                    vals = [x for x in row.values if pd.notna(x) and str(x).strip() != '']
                    if len(vals) >= 3:
                        r_id = str(vals[1]).strip()
                        val_str = vals[2]
                        raw_unit = str(vals[3]).strip().upper() if len(vals) > 3 else ''
                        val = self._parse_numeric(val_str, None)
                        if val is None: continue
                        multiplier = 1.0
                        new_unit = raw_unit
                        if raw_unit in ('KGCO2', 'KG CO2'): 
                            multiplier = 1 / 1000.0
                            new_unit = 'TCO2'
                        elif raw_unit in ('KWH', 'KW H'): 
                            multiplier = 1 / 1000.0
                            new_unit = 'MWH'
                        elif raw_unit == 'GJ': 
                            multiplier = 1 / 3.6
                            new_unit = 'MWH'
                        elif raw_unit in ('BARREL', 'BBL', 'BOE'): 
                            multiplier = 1.70  # approx MWh per barrel of crude/products
                            new_unit = 'MWH'
                        elif raw_unit in ('MBBL', 'M BBL'): 
                            multiplier = 1.70 * 1000.0
                            new_unit = 'MWH'
                        elif raw_unit == 'MMBTU': 
                            multiplier = 0.293
                            new_unit = 'MWH'
                        elif raw_unit == 'THERM': 
                            multiplier = 0.0293
                            new_unit = 'MWH'
                        
                        # Update the resource object unit to ensure metadata consistency across the model
                        if r_id in resources_dict:
                            resources_dict[r_id].unit = new_unit

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

    def _parse_time_series(self, scenario_id, resources_dict, active_sc_name, sc_name_map, years_list, unit_conversions=None):
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
                    except (ValueError, TypeError): in_target_block = False
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
                    except (ValueError, TypeError): pass
                if year is not None and start_idx != -1:
                    for i in range(start_idx + 1, len(row.values) - 1, 3):
                        r_id = str(row.values[i]).strip()
                        if r_id and r_id != 'nan' and pd.notna(row.values[i+1]):
                            price_val = self._parse_numeric(row.values[i+1], row.values[i+1])
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
                    except (ValueError, TypeError): in_target_block = False
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
                    except (ValueError, TypeError): pass
                if year is not None and start_idx != -1:
                    if len(row.values) > start_idx + 1 and pd.notna(row.values[start_idx+1]):
                        raw_prices[year] = self._parse_numeric(row.values[start_idx+1], row.values[start_idx+1])
                    if len(row.values) > start_idx + 2 and pd.notna(row.values[start_idx+2]):
                        raw_penalties[year] = self._parse_numeric(row.values[start_idx+2], 0.0)
                    if len(row.values) > start_idx + 3 and pd.notna(row.values[start_idx+3]):
                        raw_free_pi[year] = self._parse_numeric(row.values[start_idx+3], row.values[start_idx+3])
                    if len(row.values) > start_idx + 4 and pd.notna(row.values[start_idx+4]):
                        raw_free_norm[year] = self._parse_numeric(row.values[start_idx+4], row.values[start_idx+4])
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
                    except (ValueError, TypeError): in_target_block_oe = False
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
                    except (ValueError, TypeError): pass
                if year is not None and start_idx != -1:
                    for i in range(start_idx + 1, len(row.values) - 1, 3):
                        r_id = str(row.values[i]).strip()
                        if r_id and r_id != 'nan' and pd.notna(row.values[i+1]):
                            em_val = self._parse_numeric(row.values[i+1], row.values[i+1])
                            if r_id not in raw_ems: raw_ems[r_id] = {}
                            raw_ems[r_id][year] = em_val
            for r_id, e_dict in raw_ems.items():
                time_series.other_emissions_factors[r_id] = self._interpolate_dict(e_dict, years_list)

        # ── POWER LIMITS sheet ──────────────────────────────────────────────
        # Parse per-resource annual limits (e.g. max MW or MWH) and store
        # them on the time_series object after unit conversion.
        if 'POWER LIMITS' in self.xl.sheet_names:
            self._parse_power_limits(
                time_series, scenario_id, active_sc_name,
                resources_dict, years_list, unit_conversions
            )

        return time_series

    # ─────────────────────────────────────────────────────────────────────────
    # POWER LIMITS sheet parser
    # ─────────────────────────────────────────────────────────────────────────

    def _parse_power_limits(
        self,
        time_series: TimeSeriesData,
        scenario_id: str,
        active_sc_name: str,
        resources_dict: Dict[str, 'Resource'],
        years_list: list,
        unit_conversions: Dict,
    ) -> None:
        """Parse the 'POWER LIMITS' sheet and populate time_series.resource_limits.

        The sheet follows the same SCENARIO / SC-DES block convention used in
        RESSOURCES_PRICE but with a **triplet column pattern** after the YEAR
        column:  [Resource ID] [Limit Value] [Unit]  repeating for N resources.

        Special handling:
          - Year column may contain typos such as "2 025" → spaces are stripped.
          - 'LINEAR INTER' sentinel values are collected as NaN anchors and
            interpolated after the full block is read.
          - Unit conversions are applied so the stored limit is expressed in
            the resource's base unit (from resources_dict[r_id].unit).
        """
        df_pl = self.xl.parse('POWER LIMITS', header=None)

        # ── Guard: ensure unit_conversions is a dict ─────────────────────────
        if unit_conversions is None:
            unit_conversions = {}

        # ── Raw anchor dictionaries: r_id -> {year: value | 'LINEAR INTER'} ─
        raw_limits: Dict[str, Dict[int, Any]] = {}
        # ── Per-resource unit read from the sheet (last encountered wins) ────
        sheet_units: Dict[str, str] = {}

        # ── Scenario-filtering state (mirrors RESSOURCES_PRICE pattern) ──────
        in_target_block = (scenario_id is None)
        found_any_scenario_block = False

        for _, row in df_pl.iterrows():
            row_vals_upper = [
                str(x).strip().upper()
                for x in row.values
                if pd.notna(x) and str(x).strip()
            ]

            # ── Detect SCENARIO START / END markers ──────────────────────────
            if 'START' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                found_any_scenario_block = True
                in_target_block = False
                continue
            if 'END' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                in_target_block = False
                continue

            # ── SC-DES row: decide if this scenario block matches ────────────
            if 'SC-DES' in row_vals_upper and found_any_scenario_block:
                raw = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                raw_upper = [s.upper() for s in raw]
                try:
                    sc_idx = raw_upper.index('SC-DES')
                    remaining = raw_upper[sc_idx + 1:]
                    in_target_block = (
                        (scenario_id is None)
                        or any(s in [scenario_id.upper(), active_sc_name] for s in remaining)
                        or ("ALL" in remaining)
                    )
                except (ValueError, TypeError):
                    in_target_block = False
                continue

            # ── Skip rows outside the active scenario block ──────────────────
            if not in_target_block:
                continue

            # ── Skip comment rows (lines starting with **) ───────────────────
            first_non_empty = next(
                (str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()),
                ''
            )
            if first_non_empty.startswith('**'):
                continue

            # ── Try to identify the YEAR column ──────────────────────────────
            # The year may contain embedded spaces ("2 025"), so we strip them
            # before attempting int conversion.
            year = None
            start_idx = -1
            for idx, cell in enumerate(row.values):
                if pd.isna(cell):
                    continue
                cleaned = str(cell).replace(' ', '').strip()
                try:
                    y = int(cleaned)
                    if 2000 <= y <= 2100:
                        year = y
                        start_idx = idx
                        break
                except (ValueError, TypeError):
                    pass

            if year is None or start_idx == -1:
                continue

            # ── Parse triplet columns: [Resource ID] [Value] [Unit] ──────────
            remaining_cells = row.values[start_idx + 1:]
            num_remaining = len(remaining_cells)

            for offset in range(0, num_remaining - 2, 3):
                r_id_raw = remaining_cells[offset]
                val_raw = remaining_cells[offset + 1]
                unit_raw = remaining_cells[offset + 2]

                # Skip completely empty triplets
                if pd.isna(r_id_raw):
                    continue

                r_id = str(r_id_raw).strip()
                if not r_id or r_id.lower() == 'nan':
                    continue

                # ── Initialise the anchor dict for this resource ─────────────
                if r_id not in raw_limits:
                    raw_limits[r_id] = {}

                # ── Detect 'LINEAR INTER' sentinel ───────────────────────────
                val_str = str(val_raw).strip().upper() if pd.notna(val_raw) else ''
                if 'LINEAR INTER' in val_str:
                    # Store the sentinel as-is; _interpolate_dict handles it
                    raw_limits[r_id][year] = 'LINEAR INTER'
                else:
                    raw_limits[r_id][year] = self._parse_numeric(val_raw, 0.0)

                # ── Record the unit from the sheet for this resource ─────────
                if pd.notna(unit_raw):
                    sheet_units[r_id] = str(unit_raw).strip().upper()

        # ── Interpolation & unit conversion ──────────────────────────────────
        for r_id, anchors in raw_limits.items():
            # Store the raw anchors for potential sensitivity re-interpolation
            numeric_anchors = {
                y: v for y, v in anchors.items()
                if not isinstance(v, str)
            }
            time_series.resource_limits_anchors[r_id] = dict(numeric_anchors)

            # Interpolate (replaces 'LINEAR INTER' with linear values)
            interpolated = self._interpolate_dict(anchors, years_list)

            # ── Determine unit conversion multiplier ─────────────────────────
            sheet_unit = sheet_units.get(r_id, '')
            base_unit = ''
            if r_id in resources_dict:
                base_unit = str(resources_dict[r_id].unit).strip().upper()

            multiplier = 1.0
            if sheet_unit and base_unit and sheet_unit != base_unit:
                # Try exact lookup in the unit_conversions dictionary
                conv_key = (sheet_unit, base_unit)
                if conv_key in unit_conversions:
                    multiplier = unit_conversions[conv_key]
                else:
                    # Graceful fallback: common MW→MWH annual conversion
                    if sheet_unit == 'MW' and base_unit == 'MWH':
                        multiplier = 8760.0
                        warnings.warn(
                            f"[POWER LIMITS] No explicit conversion for "
                            f"{sheet_unit}→{base_unit} on '{r_id}'. "
                            f"Using standard 8760 h/year multiplier."
                        )
                    elif sheet_unit == 'MWH' and base_unit == 'MW':
                        multiplier = 1.0 / 8760.0
                        warnings.warn(
                            f"[POWER LIMITS] No explicit conversion for "
                            f"{sheet_unit}→{base_unit} on '{r_id}'. "
                            f"Using inverse 8760 h/year multiplier."
                        )
                    else:
                        warnings.warn(
                            f"[POWER LIMITS] Unknown conversion "
                            f"{sheet_unit}→{base_unit} for '{r_id}'. "
                            f"Falling back to multiplier=1.0."
                        )

            # ── Apply multiplier and store the final time series ─────────────
            time_series.resource_limits[r_id] = {
                y: val * multiplier for y, val in interpolated.items()
            }

    def _parse_public_aids(self):
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
                            active = self._parse_bool(row_list[active_idx+1])
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
                                    grant_params.rate = self._parse_numeric(row_list[grant_idx + 1], 0.0)
                                if len(row_list) > grant_idx + 2 and pd.notna(row_list[grant_idx + 2]):
                                    grant_params.cap = self._parse_numeric(row_list[grant_idx + 2], 0.0)
                                if len(row_list) > grant_idx + 3 and pd.notna(row_list[grant_idx + 3]):
                                    grant_params.renew_time = self._parse_numeric(row_list[grant_idx + 3], 0.0)
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
                                    ccfd_params.duration = int(self._parse_numeric(row_list[ccfd_idx + 1], 0))
                                if len(row_list) > ccfd_idx + 2 and pd.notna(row_list[ccfd_idx + 2]):
                                    ccfd_params.contract_type = int(self._parse_numeric(row_list[ccfd_idx + 2], 0))
                                if len(row_list) > ccfd_idx + 3 and pd.notna(row_list[ccfd_idx + 3]):
                                    ccfd_params.eua_price_pct = self._parse_numeric(row_list[ccfd_idx + 3], 0.0)
                                if len(row_list) > ccfd_idx + 4 and pd.notna(row_list[ccfd_idx + 4]):
                                    ccfd_params.nb_contracts = int(self._parse_numeric(row_list[ccfd_idx + 4], 0))
                                ccfd_params.active = True
        return grant_params, ccfd_params

    def _parse_bank_loans(self, duration):
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
                    rate = self._parse_numeric(row.get('RATE (%)', '0'), 0.0)
                    duration_val_raw = str(row.get('LOAN PERIOD (YEARS)', '1')).strip().upper()
                    if duration_val_raw == 'ALL':
                        for d in range(1, duration + 1):
                            bank_loans.append(BankLoan(rate=rate, duration=d))
                    else:
                        loan_duration = int(self._parse_numeric(duration_val_raw, 1))
                        if loan_duration < 1: loan_duration = 1
                        bank_loans.append(BankLoan(rate=rate, duration=loan_duration))
        return bank_loans

    def _parse_dac_and_credits(self, scenario_id, active_sc_name, sc_name_map, years_list):
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
                        idx = next((i for i, x in enumerate(row_list) if str(x).strip().upper() == 'ACT'), -1)
                        if idx != -1:
                            row_sc = str(row_list[idx-1]).strip().upper() if idx > 0 and pd.notna(row_list[idx-1]) else ''
                            if scenario_id and row_sc and row_sc not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]: continue
                            if len(row_list) > idx + 1 and pd.notna(row_list[idx+1]): dac_params.active = self._parse_bool(row_list[idx+1])
                            if len(row_list) > idx + 2 and pd.notna(row_list[idx+2]):
                                dac_params.start_year = int(self._parse_numeric(row_list[idx+2], dac_params.start_year))
                            if len(row_list) > idx + 3 and pd.notna(row_list[idx+3]):
                                dac_params.end_year = int(self._parse_numeric(row_list[idx+3], dac_params.end_year))
                    if 'CARAC' in row_vals:
                        vals = [x for x in row.values if pd.notna(x)]
                        if len(vals) >= 7 and str(vals[0]).strip().upper() == 'CARAC':
                            carac_val1 = str(vals[1]).strip()
                            carac_scenario = None
                            carac_start_idx = 1
                            try: float(carac_val1.replace(',','.').replace('%',''))
                            except (ValueError, TypeError):
                                carac_scenario = carac_val1.upper()
                                carac_start_idx = 2
                            if scenario_id and carac_scenario and carac_scenario not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]: continue
                            carac_vals = vals[carac_start_idx:]
                            if len(carac_vals) < 6: continue
                            year = int(self._parse_numeric(carac_vals[0], 0))
                            if year > 0:
                                raw_dac_capex[year] = self._parse_numeric(carac_vals[1], 0.0)
                                raw_dac_opex_pct[year] = self._parse_numeric(carac_vals[4], 0.0)
                                raw_dac_elec[year] = self._parse_numeric(carac_vals[5], 0.0)
                if in_credit:
                    if 'ACT' in row_vals:
                        row_list = list(row.values)
                        idx = next((i for i, x in enumerate(row_list) if str(x).strip().upper() == 'ACT'), -1)
                        if idx != -1:
                            row_sc = str(row_list[idx-1]).strip().upper() if idx > 0 and pd.notna(row_list[idx-1]) else ''
                            if scenario_id and row_sc and row_sc not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]: continue
                            if len(row_list) > idx + 1 and pd.notna(row_list[idx+1]): credit_params.active = self._parse_bool(row_list[idx+1])
                            if len(row_list) > idx + 2 and pd.notna(row_list[idx+2]):
                                credit_params.start_year = int(self._parse_numeric(row_list[idx+2], credit_params.start_year))
                            if len(row_list) > idx + 3 and pd.notna(row_list[idx+3]):
                                credit_params.end_year = int(self._parse_numeric(row_list[idx+3], credit_params.end_year))
                    if 'CREDIT' in row_vals and 'START' not in row_vals and 'END' not in row_vals:
                        vals = [x for x in row.values if pd.notna(x)]
                        if len(vals) >= 5 and str(vals[0]).strip().upper() == 'CREDIT':
                            credit_sc = str(vals[2]).strip().upper() if len(vals) > 2 else ''
                            if scenario_id and credit_sc and credit_sc not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]: continue
                            if len(vals) >= 6:
                                year = int(self._parse_numeric(vals[4], 0))
                                cost = self._parse_numeric(vals[5], 0.0)
                                if year > 0: raw_credit_cost[year] = cost
                            elif len(vals) >= 5:
                                year = int(self._parse_numeric(vals[3], 0))
                                cost = self._parse_numeric(vals[4], 0.0)
                                if year > 0: raw_credit_cost[year] = cost
                if 'TREHS' in row_vals:
                    vals = [x for x in row.values if pd.notna(x)]
                    if len(vals) >= 3 and str(vals[0]).strip().upper() == 'TREHS':
                        v_sc = None
                        v_year = 2025
                        v_pct = 1.0
                        potential_pcts = []
                        potential_years = []
                        potential_scs = []
                        for v in vals[1:]:
                            v_str = str(v).strip()
                            if '%' in v_str:
                                potential_pcts.append(self._parse_numeric(v_str, 1.0))
                                continue
                            try:
                                fv = float(v_str.replace(',','.'))
                                if fv >= 1900 and fv <= 2100: potential_years.append(int(fv))
                                else: potential_pcts.append(fv if fv <= 1.0 else fv / 100.0)
                            except: pass
                            if v_str.upper() in [s.upper() for s in sc_name_map.keys()] or v_str.upper() in [s.upper() for s in sc_name_map.values()] or v_str.upper() in ["ALL", "DEFAULT"]:
                                potential_scs.append(v_str.upper())
                        if potential_years: v_year = potential_years[0]
                        if potential_pcts: v_pct = potential_pcts[0]
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
                    dac_params.opex_by_year = {y: dac_params.capex_by_year.get(y, 0.0) * interp_opex_pct.get(y, 0.0) for y in dac_params.capex_by_year}
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
        
        time_series = self._parse_time_series(scenario_id, resources_dict, active_sc_name, sc_name_map, years_list, unit_conversions)
        
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
            raw_vals = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
            vals_upper = [v.upper() for v in raw_vals]

            if 'SENSITIVITY' in vals_upper and 'START' in vals_upper:
                in_block = True
                continue
            if 'SENSITIVITY' in vals_upper and 'END' in vals_upper:
                break

            if not in_block or not raw_vals or raw_vals[0].startswith('**'):
                continue

            tag = vals_upper[0]

            if tag == 'RUN':
                for token in raw_vals[1:]:
                    if str(token).strip().upper() in ('YES', 'NO'):
                        run = self._parse_bool(token)
                        break

            elif tag == 'VAR':
                for token in raw_vals[1:]:
                    val = self._parse_numeric(token, None)
                    if val is not None and val > 0:
                        variations.append(round(val, 6))

            elif tag == 'P/N':
                for token in raw_vals[1:]:
                    t_up = token.strip().upper()
                    if t_up in ('P', 'N', 'ALL'):
                        direction = t_up
                        break

            elif tag == 'SIM':
                placeholder_skip = ('SC1', 'SC2', 'SC3', 'SC4', 'SC...', 'SC…')
                for token in raw_vals[1:]:
                    if token.upper() not in placeholder_skip:
                        scenarios.append(token.upper())

            elif tag == 'TIME' or 'TIME' in tag:
                for token in raw_vals[1:]:
                    time_limit = int(self._parse_numeric(token, 10))
                    if time_limit > 0: break

            elif tag == 'DATA?':
                if len(raw_vals) >= 3:
                    param_name = raw_vals[1].strip()
                    targets[param_name] = self._parse_bool(raw_vals[2])
                elif len(raw_vals) == 2:
                    targets[raw_vals[1].strip()] = False

            elif tag == 'INDI':
                if len(raw_vals) >= 2:
                    indi_name = raw_vals[1].strip()
                    if indi_name and indi_name.upper() not in ('NOM', 'NAME', 'INDICATOR'):
                        indicators.append(indi_name)

        return SensitivityParams(
            run=run, variations=variations, direction=direction,
            scenarios=scenarios, time_limit=time_limit, targets=targets, indicators=indicators
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
