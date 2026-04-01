import pandas as pd
import numpy as np
import warnings
from rich import print
from tqdm import tqdm
from typing import Dict, Any
from core.model import Parameters, Resource, Technology, TimeSeriesData, EntityState, PathFinderData, Objective, Process, GrantParams, CCfDParams, BankLoan, DACParams, CreditParams, ReportingToggles

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
pd.set_option('future.no_silent_downcasting', True)

class PathFinderParser:
    def __init__(self, file_path: str, verbose: bool = False):
        self.file_path = file_path
        self.xl = pd.ExcelFile(file_path)
        self.verbose = verbose

    def _parse_scenarios(self) -> list:
        """Parse the MODELING START/END block from OverView and return list of {id, name} dicts."""
        df_overview = self.xl.parse('OverView', header=None)
        scenarios = []
        in_modeling = False
        for _, row in df_overview.iterrows():
            vals_upper = [str(x).strip().upper() for x in row.values if pd.notna(x) and str(x).strip()]
            if 'MODELING' in vals_upper and 'START' in vals_upper:
                in_modeling = True
                continue
            if 'MODELING' in vals_upper and 'END' in vals_upper:
                break
            if in_modeling and 'SC-DES' in vals_upper:
                raw = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                # Skip lines starting with '**' (usually headers or comments)
                if raw and raw[0].startswith('**'):
                    continue
                # raw[0] is 'SC-DES', raw[1] is name, raw[2] is ID
                if len(raw) >= 3:
                    sc_name = raw[1]
                    sc_id   = raw[2]
                    scenarios.append({'id': sc_id, 'name': sc_name})
        return scenarios

    def _find_blocks(self, df: pd.DataFrame) -> list:
        """Find START and END blocks in a dataframe to isolate tables."""
        blocks = []
        for i, row in tqdm(df.iterrows(), total=len(df), desc="Finding blocks", leave=False, disable=not self.verbose):
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

    def parse(self, scenario_id: str = None) -> PathFinderData:
        # Load scenario metadata once for name-based matching
        sc_meta = []
        try:
            sc_meta = self._parse_scenarios()
        except:
            pass
        sc_name_map = {s['id'].upper(): s['name'].upper() for s in sc_meta}
        active_sc_name = sc_name_map.get(scenario_id.upper() if scenario_id else "", "")

        # Load sheets
        df_overview = self.xl.parse('OverView', header=None)
        
        # 1. Parse Overview
        blocks_overview = self._find_blocks(df_overview)
        
        # Parse CHARTS toggles
        reporting_toggles = ReportingToggles()
        charts_start = next((b['row'] for b in blocks_overview if b['type'] == 'START' and b['prefix'] == 'CHARTS'), None)
        charts_end = next((b['row'] for b in blocks_overview if b['type'] == 'END' and b['prefix'] == 'CHARTS'), None)
        
        if charts_start is not None and charts_end is not None:
            df_charts = self._extract_block_data(df_overview, charts_start, charts_end)
            if not df_charts.empty:
                for _, row in df_charts.iterrows():
                    # The first column might contain chart names, second column YES/NO
                    row_vals = [str(x).strip().upper() for x in row.values if pd.notna(x) and str(x).strip()]
                    if not row_vals or row_vals[0].startswith('**'):
                        continue
                        
                    name = row_vals[0]
                    # 1. Handle Boolean Toggles (YES/NO)
                    if 'YES' in row_vals or 'NO' in row_vals:
                        is_yes = 'YES' in row_vals
                        if "EXCEL DATA" in name: reporting_toggles.results_excel = is_yes
                        elif "ENERGY MIX" in name: reporting_toggles.chart_energy_mix = is_yes
                        elif "CO2 TRAJECTORY" in name: reporting_toggles.chart_co2_trajectory = is_yes
                        elif "INDIRECT EMISSIONS" in name: reporting_toggles.chart_indirect_emissions = is_yes
                        elif "INVESTMENT PLAN" in name or "INVESTMENT COSTS" in name: reporting_toggles.chart_investment_costs = is_yes
                        elif "RESSOURCES OPEX" in name or "RESOURCE OPEX" in name: reporting_toggles.chart_resource_opex = is_yes
                        elif "CARBON TAX" in name: reporting_toggles.chart_carbon_tax_avoided = is_yes
                        elif "EXTERNAL FINANCING" in name: reporting_toggles.chart_external_financing = is_yes
                        elif "TRANSITION COST" in name: reporting_toggles.chart_transition_costs = is_yes
                        elif "CARBON PRICE" in name: reporting_toggles.chart_carbon_prices = is_yes
                        elif "SIMULATION PRICES" in name or "RESOURCE PRICES" in name: reporting_toggles.chart_resource_prices = is_yes
                        elif "INTEREST PAID" in name: reporting_toggles.chart_interest_paid = is_yes
                    
                    # 2. Handle Numeric Settings (e.g., INVESTMENT CAP)
                    elif "CAP" in name:
                        for val in row_vals[1:]:
                            try:
                                clean_val = val.replace('M€', '').replace('M', '').replace(' ', '').replace(',', '.')
                                reporting_toggles.investment_cap = float(clean_val)
                                break
                            except ValueError:
                                continue
        
        # Simplified parameters extraction (hardcoded positions based on structure):
        # Searching the whole sheet for specific keywords
        start_year = 2025
        duration = 25
        time_limit = 60.0
        mip_gap = 0.90
        relax_integrality = False
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
                    
        # Define years list for interpolation
        years_list = list(range(start_year, start_year + duration + 1))
                    
        # Find CLUSTER START -> END to get entities
        cluster_start = next((b['row'] for b in blocks_overview if b['type'] == 'START' and b['prefix'] == 'CLUSTER'), None)
        cluster_end = next((b['row'] for b in blocks_overview if b['type'] == 'END' and b['prefix'] == 'CLUSTER'), None)
        
        entities_info = {}
        if cluster_start is not None and cluster_end is not None:
            df_cluster = self._extract_block_data(df_overview, cluster_start, cluster_end)
            if 'ID' in df_cluster.columns:
                for _, row in df_cluster.iterrows():
                    e_id = str(row['ID']).strip()
                    if e_id and e_id != 'nan':
                        prod = row.get('PRODUCTION', 0.0)
                        try: prod = float(prod)
                        except: prod = 0.0
                        entities_info[e_id] = {'production': prod}
                entities = list(entities_info.keys())
                
        # Find RESOURCES
        data_start = next((b['row'] for b in blocks_overview if b['type'] == 'START' and b['prefix'] == 'DATA'), None)
        data_end = next((b['row'] for b in blocks_overview if b['type'] == 'END' and b['prefix'] == 'DATA'), None)
        
        resources_dict = {}
        if data_start is not None and data_end is not None:
            df_data = self._extract_block_data(df_overview, data_start, data_end)
            # Find the header row in block by searching for 'ID'
            for idx, row in df_data.iterrows():
                row_list = row.tolist()
                if 'ID' in row_list:
                    df_data.columns = row_list
                    df_data = df_data.iloc[idx+1:]
                    break
            
            if 'ID' in df_data.columns and 'TYPE' in df_data.columns and 'UNIT' in df_data.columns:
                for _, row in df_data.iterrows():
                    res_id = str(row['ID']).strip()
                    if res_id and pd.notna(res_id) and res_id != 'nan':
                        # Try to get a human-readable name from a NAME column, fall back to ID
                        res_name = str(row.get('NAME', '')).strip()
                        if not res_name or res_name == 'nan':
                            res_name = str(row.get('LIBELLE', '')).strip()
                        if not res_name or res_name == 'nan':
                            res_name = res_id  # fall back to ID
                        resources_dict[res_id] = Resource(
                            id=res_id,
                            type=str(row['TYPE']),
                            unit=str(row['UNIT']),
                            name=res_name
                        )
        
        # Parse OBJECTIVES
        objectives_list = []
        obj_start = next((b['row'] for b in blocks_overview if b['type'] == 'START' and b['prefix'] == 'OBJECTIVES'), None)
        obj_end = next((b['row'] for b in blocks_overview if b['type'] == 'END' and b['prefix'] == 'OBJECTIVES'), None)
        
        if obj_start is not None and obj_end is not None:
            df_obj = self._extract_block_data(df_overview, obj_start, obj_end)
            if not df_obj.empty:
                # rename columns to uniform strings
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
                df_obj = df_obj.rename(columns=col_mapping)
                
                for _, row in df_obj.iterrows():
                    res = str(row.get('RESOURCE', '')).strip()
                    if res and res != 'nan':
                        lim = str(row.get('LIMIT', 'CAP')).strip().upper()
                        # handle MIN/MAX/CAP inside strings if long descriptive limit definitions exist
                        if 'CAP' in lim: lim = 'CAP'
                        elif 'MIN' in lim: lim = 'MIN'
                        elif 'MAX' in lim: lim = 'MAX'
                        else: lim = 'CAP'
                        
                        try: t_year = int(row.get('TARGET_YEAR', 0))
                        except: t_year = 0
                        
                        try: c_val = float(row.get('CAP_VALUE', 0.0))
                        except: c_val = 0.0
                        
                        c_year = None
                        try: 
                            if pd.notna(row.get('COMP_YEAR')):
                                c_year = int(row.get('COMP_YEAR', 0))
                        except: pass
                        
                        ent = str(row.get('ENTITY', 'ALL')).strip()
                        
                        mode = str(row.get('INTERPOLATION', 'NONE')).strip().upper()
                        if mode not in ['LINEAR', 'NONE']:
                            mode = 'NONE'
                            
                        group = str(row.get('GROUP', '')).strip()
                        
                        objectives_list.append(Objective(
                            entity=ent,
                            resource=res,
                            limit_type=lim,
                            target_year=t_year,
                            cap_value=c_val,
                            comparison_year=c_year,
                            mode=mode,
                            group=group,
                            name=str(row.get('NAME', '')).strip()
                        ))
                    
        params = Parameters(
            start_year=start_year, 
            duration=duration, 
            entities=entities, 
            resources=list(resources_dict.keys()),
            time_limit=time_limit,
            mip_gap=mip_gap,
            relax_integrality=relax_integrality
        )
        
        # 2. Parse NEW TECH
        df_tech = self.xl.parse('NEW TECH', header=None)
        blocks_tech = self._find_blocks(df_tech)
        
        technologies_dict = {}
        
        tecs_start = next((b['row'] for b in blocks_tech if b['type'] == 'START' and b['prefix'] == 'TECS'), None)
        tecs_end = next((b['row'] for b in blocks_tech if b['type'] == 'END' and b['prefix'] == 'TECS'), None)
        
        if tecs_start is not None and tecs_end is not None:
            df_tecs = self._extract_block_data(df_tech, tecs_start, tecs_end)
            # find header row
            for idx, row in df_tecs.iterrows():
                row_list = row.tolist()
                if 'ID' in row_list and 'NAME' in row_list:
                    df_tecs.columns = row_list
                    df_tecs = df_tecs.iloc[idx+1:]
                    break
            
            for _, row in df_tecs.iterrows():
                t_id = str(row.get('ID', '')).strip()
                if t_id and t_id != 'nan':
                    imp_time_raw = row.get('IMPLEMANTATION TIME (YEAR)') or row.get('IMPLEMENTATION TIME (YEAR)')
                    imp_time = int(imp_time_raw) if pd.notna(imp_time_raw) else 1
                    t_name = str(row.get('NAME', '')).strip()
                    if not t_name or t_name == 'nan':
                        t_name = t_id  # fallback to ID
                    technologies_dict[t_id] = Technology(
                        id=t_id,
                        name=t_name,
                        implementation_time=imp_time,
                        capex=0.0,
                        opex=0.0,
                        impacts={}
                    )
        
        # parse COMPATIBILITIES (formerly SPECS) for CAPEX / OPEX
        euro_start = next((b['row'] for b in blocks_tech if b['type'] == 'START' and ('SPECS' in b['prefix'] or 'COMPATIBILITIES' in b['prefix']) and 'TECHNICAL' not in b['prefix']), None)
        euro_end = next((b['row'] for b in blocks_tech if b['type'] == 'END' and ('SPECS' in b['prefix'] or 'COMPATIBILITIES' in b['prefix']) and 'TECHNICAL' not in b['prefix']), None)
        
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
                    # Filter by scenario: if a SCENARIO column exists and scenario_id is set
                    row_scenario = str(row.get('SCENARIO', '')).strip().upper()
                    if scenario_id and row_scenario and row_scenario not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]:
                        # Deep check: Maybe the user used the scenario NAME instead of ID?
                        if self.verbose:
                            if self.verbose:
                                print(f"  [cyan][Ingestion][/cyan] [DEBUG] Skipping cost row for {t_id}: row scenario '{row_scenario}' != active scenario '{scenario_id.upper()}' or '{active_sc_name}'")
                        continue
                    
                    exp_type = str(row.get('TYPE (VARIABLE/FIXED)', '')).strip().upper()
                    try:
                        cost_val = float(row.get('COST', 0))
                    except:
                        cost_val = 0.0
                    
                    per_unit_str = str(row.get('PER UNIT ?', 'NO')).strip().upper()
                    is_per_unit = per_unit_str == 'YES'
                    unit_str = str(row.get('UNIT', '')).strip()

                    if exp_type == 'FIXED':
                        # CAPEX
                        if technologies_dict[t_id].capex == 0.0:
                            technologies_dict[t_id].capex = cost_val
                            technologies_dict[t_id].capex_per_unit = is_per_unit
                            technologies_dict[t_id].capex_unit = unit_str
                    elif exp_type == 'VARIABLE':
                        # OPEX
                        if technologies_dict[t_id].opex == 0.0:
                            technologies_dict[t_id].opex = cost_val
                            technologies_dict[t_id].opex_per_unit = is_per_unit
                            technologies_dict[t_id].opex_unit = unit_str
        
        # Parse technical specs for impacts
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
                    imp_type = str(row.get('TYPE (NEW/VARIATION)', '')).strip().lower() # variation or new
                    val = row.get('VALUE', 0.0)
                    try:
                        val = float(val)
                    except:
                        val = 0.0
                    
                    # Try to read the reference resource column (e.g. EN_FUEL in "0.23 m3 per GJ of EN_FUEL")
                    # Several possible column names are tried; fall back to scanning row values
                    ref_resource = None
                    for col_candidate in ['RESSOURCE REF', 'RESOURCE REF', 'REF RESSOURCE', 'REF RESOURCE',
                                          'RESSOURCE_REF', 'RESOURCE_REF', 'REF', 'REFERENCE RESSOURCE',
                                          'REFERENCE RESOURCE']:
                        col_val = row.get(col_candidate, None)
                        if col_val is not None and str(col_val).strip() not in ('', 'nan'):
                            ref_resource = str(col_val).strip()
                            break
                    
                    if ref_resource is None:
                        # Fallback: scan the row for a value that looks like a known resource ID
                        row_vals = [str(v).strip() for v in row.values if pd.notna(v)]
                        for v in row_vals:
                            if v in resources_dict and v != res_id:
                                ref_resource = v
                                break
                    
                    technologies_dict[t_id].impacts[res_id] = {
                        'type': imp_type,
                        'value': val,
                        'reference': str(row.get('STATE: INITIAL/ACTUAL/AVOIDED', 'INITIAL')).strip().upper(),
                        'ref_resource': ref_resource  # e.g. 'EN_FUEL' - the unit denominator for 'new' type
                    }
        
        if self.verbose:
            if self.verbose:
                print(f"  [cyan][Ingestion][/cyan] [TECH] Parsed {len(technologies_dict)} technologies")
        
        # 1.6 Parse Technology Compatibility Matrix
        tech_compatibilities = {}
        for i, row in df_tech.iterrows():
            row_vals = [str(x).strip().upper() for x in row if pd.notna(x)]
            if 'COMPATIBILITIES' in row_vals:
                # If this row is a START/END marker, skip it to avoid clashing with the cost block
                if any(marker in row_vals for marker in ['START', 'END']):
                    continue
                # The next row usually contains the headers (Tech IDs) starting after '**'
                header_row = df_tech.iloc[i + 1]
                headers = []
                headers_start_col = -1
                for j, cell in enumerate(header_row):
                    val = str(cell).strip()
                    if val in technologies_dict:
                        headers.append(val)
                        if headers_start_col == -1: headers_start_col = j
                
                # Matrix rows follow: [Prefix] [TechID] [x] [ ] ...
                if headers_start_col != -1:
                    for k in range(i + 2, i + 2 + len(headers)):
                        if k >= len(df_tech): break
                        row_data = df_tech.iloc[k]
                        t_id_row = str(row_data.iloc[headers_start_col - 1]).strip()
                        if t_id_row in technologies_dict:
                            compat_list = []
                            for m, h_id in enumerate(headers):
                                cell_val = str(row_data.iloc[headers_start_col + m]).strip().lower()
                                if cell_val == 'x':
                                    compat_list.append(h_id)
                            tech_compatibilities[t_id_row] = compat_list
                break
        if tech_compatibilities:
            if self.verbose:
                print(f"  [cyan][Ingestion][/cyan] [LINK] Parsed compatibility matrix for {len(tech_compatibilities)} technologies")
        
        
        # 3. Parse Entities (REFINERY)
        # Assuming the 'REFINERY' sheet or similar sheet matches the entity IDs
        # For simplicity, we loop over known entity IDs to see if there's a sheet for it.
        entities_dict = {}
        for entity_id in entities:
            # check if a sheet exists containing that ID or just try parsing "REFINERY" for TOT1
            # In our case, the overview says TOT1 is REFINERY 1. But the sheet name is 'REFINERY'.
            # We'll parse 'REFINERY' for demonstration and map it to the first entity.
            sheet_to_parse = 'REFINERY'
            if sheet_to_parse in self.xl.sheet_names:
                df_ent = self.xl.parse(sheet_to_parse, header=None)
                blocks_ent = self._find_blocks(df_ent)
                
                # TOTAL block for basic prod info
                tot_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and b['prefix'] == 'TOTAL'), None)
                tot_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and b['prefix'] == 'TOTAL'), None)
                
                production_level = entities_info.get(entity_id, {}).get('production', 0.0)
                # Parse INIT for operating hours
                init_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and b['prefix'] == 'INIT'), None)
                init_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and b['prefix'] == 'INIT'), None)
                annual_operating_hours = 8760.0
                if init_start is not None and init_end is not None:
                    df_init = df_ent.iloc[init_start+1:init_end]
                    tipe_op = 365.0
                    hours_per_day = 24.0
                    sv_act_mode = "PI"
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
                
                # Parse PROCESS block
                process_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and 'PROCESS' in b['prefix']), None)
                process_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and 'PROCESS' in b['prefix']), None)
                
                processes_dict = {}
                if process_start is not None and process_end is not None:
                    df_proc = df_ent.iloc[process_start+1:process_end]
                    for _, row in df_proc.iterrows():
                        row_vals = [str(x).strip().upper() for x in row.values if pd.notna(x)]
                        p_id = ''
                        # Basic heuristic: ID is usually right after PROCESS NAME or in the third column
                        # But simpler: look for R[number] or ID known
                        for v in row_vals:
                            if v.startswith('R') or v == 'R_OTHER':
                                p_id = v
                                break
                        
                        if p_id:
                            # Try to extract process name: first non-empty, non-ID cell before the ID
                            p_name = p_id  # default to ID
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
                                    if cell_str == 'CO2_EM':
                                        try: 
                                            v = float(row_list[i+1])
                                            p.emission_shares['CO2_EM'] = v if not np.isnan(v) else 0.0
                                        except: pass
                                    elif cell_str in resources_dict:
                                        try: 
                                            v = float(row_list[i+1])
                                            p.consumption_shares[cell_str] = v if not np.isnan(v) else 0.0
                                        except: pass
                
                # Parse BUDGET limit (CA Percentage) and TECH mapping
                ca_percent = 0.0
                tech_trans_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and b['prefix'] == 'TECHNOLOGICAL TRANSITION'), None)
                tech_trans_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and b['prefix'] == 'TECHNOLOGICAL TRANSITION'), None)
                if tech_trans_start is not None and tech_trans_end is not None:
                    df_tech_trans = df_ent.iloc[tech_trans_start+1:tech_trans_end]
                    for idx, row in df_tech_trans.iterrows():
                        row_vals = [str(x).strip().upper() for x in row.values if pd.notna(x)]
                        if 'BUDGET' in row_vals and 'CA' in row_vals:
                            # Scenario-aware budget:
                            row_vals_raw = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                            has_budget = 'BUDGET' in row_vals
                            if has_budget and scenario_id:
                                # Check if scenario ID or NAME appears in this row, or if it says 'ALL'
                                sc_match = any(v.upper() in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"] for v in row_vals_raw)
                                if not sc_match:
                                    # Skip this budget row — it belongs to another scenario
                                    continue
                                else:
                                    if self.verbose:
                                        if self.verbose:
                                            print(f"  [cyan][Ingestion][/cyan] [DEBUG] Found budget match for scenario '{scenario_id.upper()}' (row is '{row_vals_raw}')")
                                    try:
                                        ca_idx = row_vals.index('CA')
                                        if len(row_vals) > ca_idx + 1:
                                            raw_ca = str(row_vals[ca_idx + 1]).strip()
                                            has_pct_sign = '%' in raw_ca
                                            ca_val = float(raw_ca.replace('%', ''))
                                            if has_pct_sign:
                                                ca_val /= 100.0
                                            elif ca_val > 1.0:
                                                ca_val /= 100.0
                                            ca_percent = ca_val
                                            if self.verbose:
                                                print(f"  [cyan][Ingestion][/cyan] [DEBUG] Scenario '{scenario_id}': Found CA budget limit {ca_percent*100:.4f}%")
                                    except Exception as e:
                                        if self.verbose:
                                            if self.verbose:
                                                print(f"  [cyan][Ingestion][/cyan] [!] Error parsing budget CA value: {e}")
                                        pass
                            elif has_budget and not scenario_id:
                                try:
                                    ca_idx = row_vals.index('CA')
                                    if len(row_vals) > ca_idx + 1:
                                        raw_ca = str(row_vals[ca_idx + 1]).strip()
                                        has_pct_sign = '%' in raw_ca
                                        ca_val = float(raw_ca.replace('%', ''))
                                        if has_pct_sign:
                                            ca_val /= 100.0
                                        elif ca_val > 1.0:
                                            ca_val /= 100.0
                                        ca_percent = ca_val
                                        if self.verbose:
                                            print(f"  [cyan][Ingestion][/cyan] [BUDGET] CA budget limit parsed: [bold green]{ca_percent*100:.4f}%[/bold green]")
                                except:
                                    pass
                                
                        # Map valid technologies to processes
                        p_id_candidates = [v for v in row_vals if v in processes_dict]
                        if p_id_candidates:
                            p_id = p_id_candidates[0]
                            p_techs = [v for v in row_vals if v in technologies_dict]
                            # Use set-based union to avoid duplicates and allow additive list
                            existing_techs = set(processes_dict[p_id].valid_technologies)
                            new_techs = set(p_techs)
                            processes_dict[p_id].valid_technologies = list(existing_techs.union(new_techs))
                                
                # Parse PURCHASES (Sold Resources)
                sold_resources = []
                purchases_start = next((b['row'] for b in blocks_ent if b['type'] == 'START' and b['prefix'] == 'PURCHASES'), None)
                purchases_end = next((b['row'] for b in blocks_ent if b['type'] == 'END' and b['prefix'] == 'PURCHASES'), None)
                if purchases_start is not None and purchases_end is not None:
                    df_purchases = df_ent.iloc[purchases_start+1:purchases_end]
                    for idx, row in df_purchases.iterrows():
                        row_vals = [str(x).strip().upper() for x in row.values if pd.notna(x)]
                        if 'SELL' in row_vals:
                            res_id = row_vals[row_vals.index('SELL') - 1] # assuming RESSOURCE ID is before BUY/SELL
                            sold_resources.append(res_id)
                
                base_cons = {}
                base_emis = 0.0
                
                # Parse REF baselines
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
                                if val in ('RESSOURCE ID', 'RESOURCE ID'):
                                    res_col_idx = i
                                elif val in ('VALOR', 'VALUE'):
                                    val_col_idx = i
                                elif val in ('YEAR', 'ANNEE', 'ANNÉE'):
                                    year_col_idx = i
                        
                        if res_col_idx != -1 and val_col_idx != -1:
                            for j in range(idx + 1, ref_end):
                                drow = df_ent.loc[j]
                                row_vals_raw = [str(x).strip().upper() for x in drow.values if pd.notna(x)]
                                
                                # Scenario filtering for REF rows
                                if scenario_id:
                                    other_scs = [s for s in sc_name_map.keys() if s.upper() in row_vals_raw and s.upper() not in [scenario_id.upper(), str(active_sc_name).upper(), "ALL", "DEFAULT"]]
                                    if other_scs:
                                        continue # Row is for another scenario
                                
                                r_id = str(drow.iloc[res_col_idx]).strip()
                                if pd.notna(drow.iloc[val_col_idx]) and r_id and r_id != 'nan':
                                    yr = 2025 # Default if no year provided
                                    if year_col_idx != -1 and pd.notna(drow.iloc[year_col_idx]):
                                        try: yr = int(float(drow.iloc[year_col_idx]))
                                        except: pass
                                    
                                    if r_id not in ref_baselines:
                                        ref_baselines[r_id] = {}
                                    try:
                                        ref_baselines[r_id][yr] = float(drow.iloc[val_col_idx])
                                    except Exception:
                                        pass
                            break
                
                
                if tot_start is not None and tot_end is not None:
                    # df_tot might not have a strong header. we'll just scan rows
                    df_tot = df_ent.iloc[tot_start+1:tot_end]
                    for idx, row in df_tot.iterrows():
                        # Find the first non-empty cell in the row to orient columns correctly
                        vals = [x for x in row.values if pd.notna(x) and str(x).strip() != '']
                        if len(vals) >= 3:
                            # vals[0]: name (e.g., 'CO2 EMISSIONS')
                            # vals[1]: id (e.g., 'CO2_EM')
                            # vals[2]: value
                            r_id = str(vals[1]).strip()
                            r_display_name = str(vals[0]).strip()
                            # Enrich resource name if it's still using the ID as fallback
                            if r_id in resources_dict and (not resources_dict[r_id].name or resources_dict[r_id].name == r_id):
                                resources_dict[r_id].name = r_display_name
                            val_str = vals[2]
                            raw_unit = str(vals[3]).strip().upper() if len(vals) > 3 else ''
                            
                            try:
                                val = float(val_str)
                            except:
                                continue
                                
                            multiplier = 1.0
                            if raw_unit == 'KGCO2':
                                multiplier = 1 / 1000.0
                            elif raw_unit == 'KWH':
                                multiplier = 1 / 1000.0
                            elif raw_unit == 'GJ':
                                multiplier = 1 / 3.6
                                
                            total_val = val * annual_production * multiplier
                            
                            if 'CO2_EM' in r_id or r_id == 'CO2_EM':
                                base_emis += total_val
                            elif r_id in resources_dict:
                                base_cons[r_id] = total_val
                                
                entities_dict[entity_id] = EntityState(
                    id=entity_id,
                    base_consumptions=base_cons,
                    base_emissions=base_emis,
                    production_level=annual_production,
                    annual_operating_hours=annual_operating_hours,
                    sv_act_mode=sv_act_mode,
                    processes=processes_dict,
                    ref_baselines=ref_baselines,
                    ca_percentage_limit=ca_percent,
                    sold_resources=sold_resources
                )
                
        if self.verbose:
            if self.verbose:
                print(f"  [cyan][Ingestion][/cyan] [ENTITY] Parsed {len(entities_dict)} entities")
        
        # 4. Parse Time Series data
        time_series = TimeSeriesData()
        
        # Helper to interpolate a dictionary {year: value}
        def interpolate_dict(data_dict: Dict[int, Any]) -> Dict[int, float]:
            if not data_dict: return {}
            # sort by year
            s = pd.Series(data_dict).sort_index()
            # Reindex to full simulation range to avoid gaps at 0
            s = s.reindex(years_list)
            return self._interpolate_linear(s).to_dict()

        # 4.1 RESSOURCES_PRICE
        if 'RESSOURCES_PRICE' in self.xl.sheet_names:
            df_prices = self.xl.parse('RESSOURCES_PRICE', header=None)
            raw_prices = {} # r_id -> {year: val}
            
            # Detect scenario blocks: "SCENARIO N START" / "SCENARIO N END" with SC-DES inside
            # If scenario_id provided, only read the matching block.
            # If not provided (backward compat), read all rows as before.
            in_target_block = (scenario_id is None)  # If no scenario filter, read everything
            found_any_scenario_block = False
            
            for _, row in df_prices.iterrows():
                row_vals_upper = [str(x).strip().upper() for x in row.values if pd.notna(x) and str(x).strip()]
                # Detect SCENARIO N START
                if 'START' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                    found_any_scenario_block = True
                    in_target_block = False  # Will be confirmed by SC-DES row
                    continue
                # Detect SCENARIO N END
                if 'END' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                    in_target_block = False
                    continue
                # Detect SC-DES row inside a block -> check if it matches our scenario
                if 'SC-DES' in row_vals_upper and found_any_scenario_block:
                    raw = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                    raw_upper = [s.upper() for s in raw]
                    try:
                        sc_idx = raw_upper.index('SC-DES')
                        # Check all cells after SC-DES for a match
                        remaining = raw_upper[sc_idx+1:]
                        in_target_block = (scenario_id is None) or any(s in [scenario_id.upper(), active_sc_name] for s in remaining) or ("ALL" in remaining)
                        if self.verbose:
                            if self.verbose:
                                print(f"  [cyan][Ingestion][/cyan] [DEBUG] Skipping price block (Scenarios in row: {remaining})")
                    except:
                        in_target_block = False
                    continue
                
                if not in_target_block:
                    continue
                    
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
                            price_val = row.values[i+1]
                            if r_id not in raw_prices:
                                raw_prices[r_id] = {}
                            raw_prices[r_id][year] = price_val
                            
            for r_id, p_dict in raw_prices.items():
                if r_id not in resources_dict:
                    # Auto-register resource if it exists in price sheet but not in DATA block
                    res_name = r_id
                    res_type = 'PRODUCTION' # Default type for buyable resources
                    res_unit = 'units'      # Default unit
                    resources_dict[r_id] = Resource(id=r_id, type=res_type, unit=res_unit, name=res_name)
                    
                time_series.resource_prices[r_id] = interpolate_dict(p_dict)
            if self.verbose:
                sc_tag = f" [scenario={scenario_id}]" if scenario_id else ""
                if self.verbose:
                    print(f"  [cyan][Ingestion][/cyan] [PRICE] Parsed resource prices for {len(time_series.resource_prices)} resources{sc_tag}")

        # 4.2 CARBON QUOTAS
        if 'CARBON QUOTAS' in self.xl.sheet_names:
            df_cq = self.xl.parse('CARBON QUOTAS', header=None)
            raw_prices = {} # {year: price}
            raw_free_pi = {} # {year: free_quota_pct}
            raw_free_norm = {} # {year: N-EAU_pct}
            raw_penalties = {} # {year: penalty_factor}

            # Scenario filtering logic similar to RESSOURCES_PRICE
            in_target_block = (scenario_id is None)
            found_any_scenario_block = False
            
            for _, row in df_cq.iterrows():
                row_vals_upper = [str(x).strip().upper() for x in row.values if pd.notna(x) and str(x).strip()]
                
                # Detect SCENARIO N START
                if 'START' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                    found_any_scenario_block = True
                    in_target_block = False
                    continue
                # Detect SCENARIO N END
                if 'END' in row_vals_upper and any('SCENARIO' in v for v in row_vals_upper):
                    in_target_block = False
                    continue
                # Detect SC-DES row
                if 'SC-DES' in row_vals_upper and found_any_scenario_block:
                    raw = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
                    raw_upper = [s.upper() for s in raw]
                    try:
                        sc_idx = raw_upper.index('SC-DES')
                        remaining = raw_upper[sc_idx+1:]
                        in_target_block = (scenario_id is None) or any(s in [scenario_id.upper(), active_sc_name] for s in remaining) or ("ALL" in remaining)
                    except:
                        in_target_block = False
                    continue

                if not in_target_block:
                    continue

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
                    # Col start_idx + 1: PRICE, Col start_idx + 2: PENALTY, Col start_idx + 3: PI-EAU, Col start_idx + 4: N-EAU
                    if len(row.values) > start_idx + 1 and pd.notna(row.values[start_idx+1]):
                        raw_prices[year] = row.values[start_idx+1]
                    if len(row.values) > start_idx + 2 and pd.notna(row.values[start_idx+2]):
                        try:
                            raw_penalties[year] = float(row.values[start_idx+2])
                        except: pass
                    if len(row.values) > start_idx + 3 and pd.notna(row.values[start_idx+3]):
                        raw_free_pi[year] = row.values[start_idx+3]
                    if len(row.values) > start_idx + 4 and pd.notna(row.values[start_idx+4]):
                        raw_free_norm[year] = row.values[start_idx+4]
                        
            time_series.carbon_prices = interpolate_dict(raw_prices)
            time_series.carbon_quotas_pi = interpolate_dict(raw_free_pi)
            time_series.carbon_quotas_norm = interpolate_dict(raw_free_norm)
            
            # --- Monotonic Penalty Price Logic ---
            interp_factors = interpolate_dict(raw_penalties)
            adjusted_factors = {}
            last_eff_price = -1.0
            
            for y in sorted(time_series.carbon_prices.keys()):
                p = time_series.carbon_prices[y]
                f = interp_factors.get(y, 0.0)
                eff_price = p * (1.0 + f)
                
                if eff_price < last_eff_price:
                    if p > 0:
                        f = (last_eff_price / p) - 1.0
                    else:
                        f = 0.0
                    eff_price = last_eff_price
                
                adjusted_factors[y] = f
                last_eff_price = eff_price
                
            time_series.carbon_penalties = adjusted_factors
            
            if self.verbose:
                print(f"  [cyan][Ingestion][/cyan] [CO2] Parsed carbon prices for {len(time_series.carbon_prices)} years")
                print(f"  [cyan][Ingestion][/cyan] [QUOTA] Parsed free PI carbon quotas for {len(time_series.carbon_quotas_pi)} years")
                print(f"  [cyan][Ingestion][/cyan] [QUOTA] Parsed free NORM carbon quotas for {len(time_series.carbon_quotas_norm)} years")
                print(f"  [cyan][Ingestion][/cyan] [!] Parsed carbon penalties for {len(time_series.carbon_penalties)} years (monotonic constraint applied)")

        # 4.3 OTHER EMISSIONS
        if 'OTHER EMISSIONS' in self.xl.sheet_names:
            df_oe = self.xl.parse('OTHER EMISSIONS', header=None)
            raw_ems = {} # r_id -> {year: val}
            
            # Same scenario block logic as RESSOURCES_PRICE
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
                        if self.verbose:
                            print(f"  [cyan][Ingestion][/cyan] [DEBUG] Skipping Other Emissions block (Scenarios in row: {remaining})")
                    except:
                        in_target_block_oe = False
                    continue
                
                if not in_target_block_oe:
                    continue
                
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
                            
                            if 'H2' in r_id.upper():
                                try:
                                    # User inputs are traditionally in kgCO2/kgH2 (e.g., Grey=10, Blue=5).
                                    # Model expects tCO2/GJ. 1 kgH2 = 0.12 GJ -> 8.333 kgH2/GJ.
                                    # em_val * 8.333 / 1000 = em_val / 120.0
                                    em_val = float(em_val) / 120.0
                                except:
                                    pass
                                
                            if r_id not in raw_ems:
                                raw_ems[r_id] = {}
                            raw_ems[r_id][year] = em_val
                            
            for r_id, e_dict in raw_ems.items():
                time_series.other_emissions_factors[r_id] = interpolate_dict(e_dict)
            if self.verbose:
                if self.verbose:
                    print(f"  [cyan][Ingestion][/cyan] [EMS] Parsed other emissions for {len(time_series.other_emissions_factors)} factors")

        # 5. Parse PUBLIC AID
        grant_params = GrantParams()
        ccfd_params = CCfDParams()
        
        if 'PUBLIC AID' in self.xl.sheet_names:
            df_aid = self.xl.parse('PUBLIC AID', header=None)
            blocks_aid = self._find_blocks(df_aid)
            
            # Check ACTIVE
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
                            # We search for GRANT in the original row list to get exact column index
                            grant_idx = -1
                            for i, x in enumerate(row_list):
                                if str(x).strip().upper() == 'GRANT':
                                    grant_idx = i
                                    break
                            
                            if grant_idx != -1:
                                if len(row_list) > grant_idx + 1 and pd.notna(row_list[grant_idx + 1]):
                                    try: 
                                        grant_params.rate = float(str(row_list[grant_idx + 1]).replace('%', '')) / 100.0 if '%' in str(row_list[grant_idx + 1]) else float(row_list[grant_idx + 1])
                                    except: pass
                                if len(row_list) > grant_idx + 2 and pd.notna(row_list[grant_idx + 2]):
                                    try: grant_params.cap = float(row_list[grant_idx + 2])
                                    except: pass
                                if len(row_list) > grant_idx + 3 and pd.notna(row_list[grant_idx + 3]):
                                    try: grant_params.renew_time = float(row_list[grant_idx + 3])
                                    except: pass
                                grant_params.active = True
                                
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
                                    try: 
                                        ccfd_params.eua_price_pct = float(str(row_list[ccfd_idx + 3]).replace('%', '')) / 100.0 if '%' in str(row_list[ccfd_idx + 3]) else float(row_list[ccfd_idx + 3])
                                    except: pass
                                if len(row_list) > ccfd_idx + 4 and pd.notna(row_list[ccfd_idx + 4]):
                                    try: ccfd_params.nb_contracts = int(float(row_list[ccfd_idx + 4]))
                                    except: pass
                                ccfd_params.active = True

        # 6. Parse BANK LOANS
        bank_loans = []
        if 'BANK' in self.xl.sheet_names:
            df_bank = self.xl.parse('BANK', header=None)
            blocks_bank = self._find_blocks(df_bank)
            
            prod_start = next((b['row'] for b in blocks_bank if b['type'] == 'START' and b['prefix'] == 'PRODUCTS'), None)
            prod_end = next((b['row'] for b in blocks_bank if b['type'] == 'END' and b['prefix'] == 'PRODUCTS'), None)
            
            if prod_start is not None and prod_end is not None:
                df_prod = self._extract_block_data(df_bank, prod_start, prod_end)
                # Find header row
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
                            for d in range(1, params.duration + 1):
                                bank_loans.append(BankLoan(rate=rate, duration=d))
                        else:
                            loan_duration = int(float(duration_val_raw))
                            if loan_duration < 1: loan_duration = 1
                            bank_loans.append(BankLoan(rate=rate, duration=loan_duration))
                    except:
                        pass
            if self.verbose:
                if self.verbose:
                    print(f"  [cyan][Ingestion][/cyan] [BANK] Parsed {len(bank_loans)} bank loan products")

        # 7. Parse DAC and Credits
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
                if not row_vals:
                    continue
                
                if ('DAC' in row_vals) or ('DIRECT AIR CAPTURE' in row_vals):
                    in_dac = True
                    dac_params.active = True
                
                if ('CREDIT' in row_vals or 'CARBON CREDIT' in row_vals) and ('START' in row_vals):
                    in_credit = True
                    credit_params.active = True
                    
                if 'DAC' in row_vals and 'END' in row_vals:
                    in_dac = False
                
                if 'CREDIT' in row_vals and 'END' in row_vals:
                    in_credit = False
                    
                if in_dac:
                    if 'ACT' in row_vals:
                        row_list = list(row.values)
                        try:
                            # 'ACT' is usually in column 1, 'YES' in col 2, Start Year col 3, End Year col 4
                            idx = next((i for i, x in enumerate(row_list) if str(x).strip().upper() == 'ACT'), -1)
                            if idx != -1:
                                # Check scenario if present before 'ACT'
                                row_sc = str(row_list[idx-1]).strip().upper() if idx > 0 and pd.notna(row_list[idx-1]) else ''
                                if scenario_id and row_sc and row_sc not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]:
                                    if self.verbose:
                                        print(f"  [cyan][Ingestion][/cyan] [DEBUG] Skipping DAC ACT row: '{row_sc}' != '{scenario_id.upper()}' or '{active_sc_name}'")
                                    continue
                                
                                if len(row_list) > idx + 1 and pd.notna(row_list[idx+1]):
                                    dac_params.active = (str(row_list[idx+1]).strip().upper() == 'YES')
                                if len(row_list) > idx + 2 and pd.notna(row_list[idx+2]):
                                    val = str(row_list[idx+2]).strip()
                                    if val and val.replace('.', '', 1).isnumeric():
                                        dac_params.start_year = int(float(val))
                                if len(row_list) > idx + 3 and pd.notna(row_list[idx+3]):
                                    val = str(row_list[idx+3]).strip()
                                    if val and val.replace('.', '', 1).isnumeric():
                                        dac_params.end_year = int(float(val))
                        except: pass
                        
                        
                    if 'CARAC' in row_vals:
                        vals = [x for x in row.values if pd.notna(x)]
                        if len(vals) >= 7 and str(vals[0]).strip().upper() == 'CARAC':
                            # CARAC rows: vals[1] may be a scenario ID (string) or a year (int)
                            # Structure: CARAC, <scenario_id>, year, capex, ..., opex_pct, ..., elec
                            # If scenario_id is provided, skip non-matching rows
                            carac_val1 = str(vals[1]).strip()
                            # Detect if vals[1] is a scenario ID (not a year)
                            carac_scenario = None
                            carac_start_idx = 1  # default: year is at index 1
                            try:
                                int(float(carac_val1))
                                # It's a year, no scenario prefix
                            except (ValueError, TypeError):
                                # It's a scenario ID
                                carac_scenario = carac_val1.upper()
                                carac_start_idx = 2
                            
                            # Skip if this CARAC row is for a different scenario
                            if scenario_id and carac_scenario and carac_scenario not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]:
                                if self.verbose:
                                    print(f"  [cyan][Ingestion][/cyan] [DEBUG] Skipping DAC CARAC row: '{carac_scenario}' != '{scenario_id.upper()}' or '{active_sc_name}'")
                                continue
                            
                            # Rebase vals if scenario prefix present
                            carac_vals = vals[carac_start_idx:]
                            if len(carac_vals) < 6:
                                continue
                            
                            try:
                                year = int(carac_vals[0])
                                
                                try: raw_dac_capex[year] = float(carac_vals[1])
                                except: raw_dac_capex[year] = str(carac_vals[1]).strip()
                                
                                try: raw_dac_opex_pct[year] = float(carac_vals[4])
                                except: raw_dac_opex_pct[year] = str(carac_vals[4]).strip()
                                
                                try: raw_dac_elec[year] = float(carac_vals[5])
                                except: raw_dac_elec[year] = str(carac_vals[5]).strip()
                            except Exception:
                                pass
                
                if in_credit:
                    if 'ACT' in row_vals:
                        row_list = list(row.values)
                        try:
                            idx = next((i for i, x in enumerate(row_list) if str(x).strip().upper() == 'ACT'), -1)
                            if idx != -1:
                                row_sc = str(row_list[idx-1]).strip().upper() if idx > 0 and pd.notna(row_list[idx-1]) else ''
                                if scenario_id and row_sc and row_sc not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]:
                                    if self.verbose:
                                        print(f"  [cyan][Ingestion][/cyan] [DEBUG] Skipping Credit ACT row: '{row_sc}' != '{scenario_id.upper()}' or '{active_sc_name}'")
                                    continue
                                    
                                if len(row_list) > idx + 1 and pd.notna(row_list[idx+1]):
                                    credit_params.active = (str(row_list[idx+1]).strip().upper() == 'YES')
                                if len(row_list) > idx + 2 and pd.notna(row_list[idx+2]):
                                    val = str(row_list[idx+2]).strip()
                                    if val and val.replace('.', '', 1).isnumeric():
                                        credit_params.start_year = int(float(val))
                                if len(row_list) > idx + 3 and pd.notna(row_list[idx+3]):
                                    val = str(row_list[idx+3]).strip()
                                    if val and val.replace('.', '', 1).isnumeric():
                                        credit_params.end_year = int(float(val))
                        except: pass
                    if 'CREDIT' in row_vals and 'START' not in row_vals and 'END' not in row_vals:
                        vals = [x for x in row.values if pd.notna(x)]
                        if len(vals) >= 5 and str(vals[0]).strip().upper() == 'CREDIT':
                            # Format: CREDIT, resource_id, scenario_id, unit, year, cost
                            credit_sc = str(vals[2]).strip().upper() if len(vals) > 2 else ''
                            if scenario_id and credit_sc and credit_sc not in [scenario_id.upper(), active_sc_name, "ALL", "DEFAULT"]:
                                # Skip row for other scenario
                                continue
                            
                            try:
                                # Snippet: CREDIT, resource_id, scenario_id, unit, year, cost
                                # vals: [CREDIT, CO2_EM, LCB, tCO2, 2025, 22.50 €]
                                if len(vals) >= 6:
                                    year = int(float(vals[4]))
                                    cost_raw = str(vals[5]).strip()
                                    # Handle European format: "22,50 €" -> 22.5
                                    cost_clean = cost_raw.replace('€', '').replace('$', '').replace(',', '.').strip()
                                    cost = float(cost_clean)
                                    raw_credit_cost[year] = cost
                                elif len(vals) >= 5:
                                    # Fallback index if unit is missing
                                    year = int(float(vals[3]))
                                    cost = float(str(vals[4]).replace('€', '').replace(',', '.').strip())
                                    raw_credit_cost[year] = cost
                            except Exception as e:
                                if self.verbose:
                                    print(f"  [cyan][Ingestion][/cyan] [DEBUG] Error parsing Credit line: {e}. vals={vals}")
                                pass
                # Common parsers for both DAC and Credit (outside specific ACT/CARAC loops)
                if 'TREHS' in row_vals:
                    vals = [x for x in row.values if pd.notna(x)]
                    if len(vals) >= 3 and str(vals[0]).strip().upper() == 'TREHS':
                        # Flexible Format: TREHS, [Scenario?], [Year?], [Pct?], ...
                        # We look for a Year (2000-2100 or 1900-2000), a Scenario ID, and a Percentage.
                        v_sc = None
                        v_year = 2025
                        v_pct = 100.0
                        
                        potential_pcts = []
                        potential_years = []
                        potential_scs = []
                        
                        for v in vals[1:]:
                            v_str = str(v).strip()
                            # Check for Percentage
                            if isinstance(v, str) and '%' in v_str:
                                try: potential_pcts.append(float(v_str.replace('%', '')) / 100.0)
                                except: pass
                                continue
                            
                            # Check for Years/Numbers
                            try:
                                fv = float(v_str)
                                if (fv >= 1900 and fv <= 2100) and v_year == 2025:
                                    potential_years.append(int(fv))
                                elif fv <= 100.0:
                                    potential_pcts.append(fv)
                                continue
                            except: pass
                            
                            # Everything else is a potential Scenario
                            if v_str.upper() in [s.upper() for s in sc_name_map.keys()] or v_str.upper() in [s.upper() for s in sc_name_map.values()] or v_str.upper() in ["ALL", "DEFAULT"]:
                                potential_scs.append(v_str.upper())
                        
                        # Resolution
                        if potential_years: v_year = potential_years[0]
                        if potential_pcts: 
                            v_pct = potential_pcts[0]
                            # If we found two numbers (Year and Pct), but Year was first:
                            if len(potential_years) > 1 and len(potential_pcts) == 1:
                                # Maybe the second year was actually the percentage? (unlikely)
                                pass
                            if v_pct > 1.0: v_pct = v_pct / 100.0
                            
                        if potential_scs: v_sc = potential_scs[0]
                            
                        if scenario_id and v_sc and v_sc not in [scenario_id.upper(), str(active_sc_name).upper(), "ALL", "DEFAULT"]:
                            continue
                            
                        if in_dac:
                            dac_params.ref_year = v_year
                            dac_params.max_volume_pct = v_pct
                        elif in_credit:
                            credit_params.ref_year = v_year
                            credit_params.max_volume_pct = v_pct

            if dac_params.active:
                dac_params.capex_by_year = interpolate_dict(raw_dac_capex)
                if not dac_params.capex_by_year:
                    dac_params.active = False
                else:
                    interp_opex_pct = interpolate_dict(raw_dac_opex_pct)
                    dac_params.opex_by_year = {y: dac_params.capex_by_year.get(y, 0.0) * (interp_opex_pct.get(y, 0.0) / 100.0) for y in dac_params.capex_by_year}
                    dac_params.elec_by_year = interpolate_dict(raw_dac_elec)
                
                if self.verbose:
                    print(f"  [cyan][Ingestion][/cyan] [DAC] Parsed DAC parameters (Active: {dac_params.active}, Cap: {dac_params.max_volume_pct*100:.1f}%, Ref Year: {dac_params.ref_year})")
                
            if credit_params.active:
                credit_params.cost_by_year = interpolate_dict(raw_credit_cost)
                if not credit_params.cost_by_year:
                    credit_params.active = False
                else:
                    if self.verbose:
                        print(f"  [cyan][Ingestion][/cyan] [CREDIT] Parsed Credit parameters (Cap: {credit_params.max_volume_pct*100}%, Ref Year: {credit_params.ref_year})")

        return PathFinderData(
            parameters=params,
            resources=resources_dict,
            technologies=technologies_dict,
            time_series=time_series,
            entities=entities_dict,
            objectives=objectives_list,
            tech_compatibilities=tech_compatibilities,
            grant_params=grant_params,
            ccfd_params=ccfd_params,
            bank_loans=bank_loans,
            dac_params=dac_params,
            credit_params=credit_params,
            reporting_toggles=reporting_toggles
        )

if __name__ == '__main__':
    parser = PathFinderParser('PathFinder input.xlsx')
    data = parser.parse()
    # print("Done")
if __name__ == '__main__':
    parser = PathFinderParser('PathFinder input.xlsx')
    data = parser.parse()
    # print("Done")
