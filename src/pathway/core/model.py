from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Any

@dataclass
class GrantParams:
    active: bool = False
    rate: float = 0.0          # e.g., 0.7 for 70% of capex
    cap: float = float('inf')  # e.g., 100M€ max
    renew_time: float = 0.0    # years before a new grant can be obtained
    excluded_technologies: List[str] = field(default_factory=list)

@dataclass
class CCfDParams:
    active: bool = False
    duration: int = 0
    contract_type: int = 2
    eua_price_pct: float = 1.0
    nb_contracts: int = 0

@dataclass
class BankLoan:
    rate: float = 0.0
    duration: int = 1

@dataclass
class DACParams:
    active: bool = False
    start_year: Optional[int] = None
    end_year: Optional[int] = None
    capex_by_year: Dict[int, float] = field(default_factory=dict) # year -> CAPEX for 1 tCO2
    opex_by_year: Dict[int, float] = field(default_factory=dict)  # year -> OPEX for 1 tCO2
    elec_by_year: Dict[int, float] = field(default_factory=dict)  # year -> kWh for 1 tCO2
    max_volume_pct: float = 1.0 # Default to 100% (unlimited)
    ref_year: int = 2025

@dataclass
class CreditParams:
    active: bool = False
    start_year: Optional[int] = None
    end_year: Optional[int] = None
    cost_by_year: Dict[int, float] = field(default_factory=dict) # year -> COST (€) for 1 tCO2
    max_volume_pct: float = 0.0 # e.g. 0.05 for 5% max
    ref_year: int = 2025 # reference year for the 5% cap

@dataclass
class ReportingToggles:
    results_excel: bool = True
    chart_energy_mix: bool = True
    chart_co2_trajectory: bool = True
    chart_indirect_emissions: bool = True
    chart_investment_costs: bool = True
    chart_total_opex: bool = True
    chart_carbon_tax_avoided: bool = True
    chart_external_financing: bool = True
    chart_transition_costs: bool = True
    chart_carbon_prices: bool = True
    chart_interest_paid: bool = True
    chart_resource_prices: bool = True
    chart_co2_abatement_cost: bool = True # New chart for Marginal Abatement Cost
    investment_cap: float = 0.0 # New field for dynamic cap on charts

@dataclass
class SensitivityParams:
    """
    Paramètres d'analyse de sensibilité extraits du bloc SENSITIVITY du fichier Excel.
    Utilisés par le script run_sensitivity.py pour piloter les simulations OAT.
    """
    # Indique si l'analyse de sensibilité doit être exécutée (RUN: YES/NO)
    run: bool = False
    # Liste des amplitudes de variation (ex: [0.05, 0.10, 0.25, 0.50, 1.00])
    variations: List[float] = field(default_factory=list)
    # Direction des variations : 'P' (positif), 'N' (négatif), ou 'ALL' (les deux)
    direction: str = "ALL"
    # Identifiants des scénarios à simuler (ex: ['BS'])
    scenarios: List[str] = field(default_factory=list)
    # Temps limite alloué au solveur par simulation (secondes)
    time_limit: int = 10
    # Dictionnaire des données à perturber : {'EUA': True, 'RESSOURCES PRICE': False, ...}
    targets: Dict[str, bool] = field(default_factory=dict)
    # Liste des indicateurs KPI à extraire (ex: ['TRANSITION COST', 'AVERAGE CO2 ABATEMENT', ...])
    indicators: List[str] = field(default_factory=list)


@dataclass
class Objective:
    entity: str
    resource: str
    limit_type: str # 'MIN', 'MAX', 'CAP'
    target_year: int
    cap_value: float = 0.0
    comparison_year: Optional[int] = None
    mode: str = 'NONE'
    group: str = ""
    name: str = ""
    penalty_type: str = "AT ALL COST"  # Options: "NONE", "AT ALL COST"

@dataclass
class Parameters:
    start_year: int
    duration: int
    entities: List[str]
    resources: List[str]
    time_limit: float = 60.0
    mip_gap: float = 0.90
    relax_integrality: bool = False
    # potentially other overview parameters like discount rate if available

@dataclass
class Resource:
    id: str
    type: str  # Consommation/Émission/Production
    unit: str
    name: str = ""  # Human-readable display name (e.g. "Fuel Gas" vs id "EN_FUEL")
    category: str = "Other"

@dataclass
class Technology:
    id: str
    name: str = ""
    implementation_time: int = 1
    capex: float = 0.0
    capex_per_unit: bool = False
    capex_unit: str = ""
    capex_by_year: Dict[int, float] = field(default_factory=dict)
    opex: float = 0.0
    opex_per_unit: bool = False
    opex_unit: str = ""
    opex_by_year: Dict[int, float] = field(default_factory=dict)
    
    # Raw values from Excel before interpolation (for sensitivity re-interpolation)
    capex_anchors: Dict[int, float] = field(default_factory=dict)
    opex_anchors: Dict[int, float] = field(default_factory=dict)
    # impacts: e.g., {'ELEC': {'type': 'variation', 'value': 0.1}, 'CO2': {'type': 'new', 'value': -50}}
    impacts: Dict[str, Dict[str, float]] = field(default_factory=dict)

@dataclass
class TimeSeriesData:
    resource_prices: Dict[str, Dict[int, float]] = field(default_factory=dict)  # resource_id -> year -> price
    carbon_quotas_pi: Dict[int, float] = field(default_factory=dict) # year -> free quota volume or % for PI
    carbon_quotas_norm: Dict[int, float] = field(default_factory=dict) # year -> free quota volume or % for NORM
    carbon_prices: Dict[int, float] = field(default_factory=dict) # year -> CO2 price
    carbon_penalties: Dict[int, float] = field(default_factory=dict) # year -> penalty factor (x in 1+x)
    other_emissions_factors: Dict[str, Dict[int, float]] = field(default_factory=dict) # resource_id -> year -> factor
    
    # Raw values from Excel before interpolation (for sensitivity re-interpolation)
    resource_prices_anchors: Dict[str, Dict[int, float]] = field(default_factory=dict)
    carbon_prices_anchors: Dict[int, float] = field(default_factory=dict)

@dataclass
class Process:
    id: str
    name: str = ""
    nb_units: int = 1
    consumption_shares: Dict[str, float] = field(default_factory=dict)
    emission_shares: Dict[str, float] = field(default_factory=dict)
    valid_technologies: List[str] = field(default_factory=list)

@dataclass
class EntityState:
    id: str
    # Initial consumptions/emissions
    base_consumptions: Dict[str, float] = field(default_factory=dict)
    base_emissions: float = 0.0
    production_level: float = 0.0
    annual_operating_hours: float = 8760.0
    sv_act_mode: str = "PI"
    processes: Dict[str, Process] = field(default_factory=dict)
    # Reference Objectives Baselines
    ref_baselines: Dict[str, Dict[int, float]] = field(default_factory=dict) # resource_id -> year -> value
    # Financial configs for CA-based budget constraints
    ca_percentage_limit: float = 0.0
    sold_resources: List[str] = field(default_factory=list)

@dataclass
class PathFinderData:
    parameters: Parameters
    resources: Dict[str, Resource]
    technologies: Dict[str, Technology]
    time_series: TimeSeriesData
    entities: Dict[str, EntityState]
    objectives: List[Objective] = field(default_factory=list)
    # Maps parent resource ID -> list of sub-type IDs (MIXT OF relationships)
    mixt_groups: Dict[str, List[str]] = field(default_factory=dict)
    # Maps technology ID -> list of compatible technology IDs
    tech_compatibilities: Dict[str, List[str]] = field(default_factory=dict)
    # Maps (unit_in, unit_out) -> conversion factor
    unit_conversions: Dict[Tuple[str, str], float] = field(default_factory=dict)
    
    grant_params: GrantParams = field(default_factory=GrantParams)
    ccfd_params: CCfDParams = field(default_factory=CCfDParams)
    
    # Financial Params
    bank_loans: List[BankLoan] = field(default_factory=list)
    
    # DAC & Credits
    dac_params: DACParams = field(default_factory=DACParams)
    credit_params: CreditParams = field(default_factory=CreditParams)
    
    reporting_toggles: ReportingToggles = field(default_factory=ReportingToggles)
