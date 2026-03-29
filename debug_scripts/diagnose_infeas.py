import pulp
from ingestion import PathFinderParser
from optimizer import PathFinderOptimizer

def diagnose_infeasibility():
    print("--- Loading data ---")
    parser = PathFinderParser('PathFinder input.xlsx')
    data = parser.parse()
    
    # Check original solver
    print("\n--- Solving Full Model ---")
    opt = PathFinderOptimizer(data)
    opt.build_model()
    # Write model to file for manual inspection if needed
    opt.model.writeLP("diagnose_infeas.lp")
    status = opt.solve()
    print(f"Status: {status}")
    
    if status == 'Infeasible':
        # Diagnostic 1: Remove Budget Limit
        print("\n--- Test 1: Full Model MINUS Budget Limit ---")
        entity = list(data.entities.values())[0]
        original_ca = entity.ca_percentage_limit
        entity.ca_percentage_limit = 0.0
        opt1 = PathFinderOptimizer(data)
        opt1.build_model()
        status1 = opt1.solve()
        print(f"Status Test 1: {status1}")
        entity.ca_percentage_limit = original_ca # Restore
        
        # Diagnostic 2: Remove Objectives
        print("\n--- Test 2: Full Model MINUS Goals/Objectives ---")
        original_objs = data.objectives
        data.objectives = []
        opt2 = PathFinderOptimizer(data)
        opt2.build_model()
        status2 = opt2.solve()
        print(f"Status Test 2: {status2}")
        data.objectives = original_objs # Restore

        # Diagnostic 3: Relax Emissions constraint
        print("\n--- Test 3: Checking CO2 targets vs Baseline ---")
        for obj in original_objs:
            if obj.resource == 'CO2_EM':
                print(f"Goal: {obj.limit_type} {obj.cap_value} at {obj.target_year} (Comp: {obj.comparison_year})")
        
        print("\n--- Detailed Slack Trace ---")
        print("Note: Infeasibility is often caused by objectives that require more Capex than the Budget allows,")
        print("or by technological impacts that cannot reach the target (e.g. 90% reduction required but max is 87.5%).")

if __name__ == "__main__":
    diagnose_infeasibility()
