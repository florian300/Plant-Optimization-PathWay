import pickle
import pandas as pd
from ingestion import PathFinderParser
from optimizer import PathFinderOptimizer

parser = PathFinderParser("PathFinder input.xlsx")
try:
    data = parser.parse()
    opt = PathFinderOptimizer(data)
    opt.build_model()
    opt.solve()

    print(f"GRANT params: {opt.data.grant_params}")
    print(f"CCFD params: {opt.data.ccfd_params}")
    
    # Analyze raw CCfD subsidies
    if hasattr(opt.data, 'ccfd_params') and opt.data.ccfd_params.active:
        print("\n--- CCfD Subsidies Simulation ---")
        ccfd_p = opt.data.ccfd_params
        for t in opt.years:
            c_price_t = opt.data.time_series.carbon_prices.get(t, 0.0)
            strike_price = (1.0 + ccfd_p.eua_price_pct) * c_price_t
            
            # Simulated ton for checking price diff
            subsidy_per_ton = 0.0
            for tau in range(t, min(t + ccfd_p.duration, opt.years[-1] + 1)):
                if tau in opt.years:
                    c_price_tau = opt.data.time_series.carbon_prices.get(tau, 0.0)
                    if ccfd_p.contract_type == 2:
                        subsidy_per_ton += (strike_price - c_price_tau)
                    else: # type == 1
                        subsidy_per_ton += max(0, strike_price - c_price_tau)
            
            print(f"Year {t} Investment - Base C_Price: {c_price_t:.2f}, Strike: {strike_price:.2f} -> Tot Subsidy/ton over 10y: {subsidy_per_ton:.2f}")

    grant_count = 0
    ccfd_count = 0

    if hasattr(opt, 'grant_used_vars'):
        for k, v in opt.grant_used_vars.items():
            if v.varValue is not None and v.varValue > 0.5:
                print(f"GRANT USED: {k} -> {v.varValue}")
                grant_count += 1
                
    if hasattr(opt, 'ccfd_used_vars'):
        for k, v in opt.ccfd_used_vars.items():
            if v.varValue is not None and v.varValue > 0.5:
                print(f"CCFD USED: {k} -> {v.varValue}")
                ccfd_count += 1
                
    print(f"TOTAL GRANTS: {grant_count}")
    print(f"TOTAL CCFDS: {ccfd_count}")
    
except Exception as e:
    import traceback
    traceback.print_exc()
