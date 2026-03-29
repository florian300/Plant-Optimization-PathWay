import pulp
from ingestion import PathFinderParser
from optimizer import PathFinderOptimizer

parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()
data.objectives = [] # Disable constraints

optimizer = PathFinderOptimizer(data)
optimizer.build_model()
optimizer.solve()

for t in optimizer.years:
    emis = optimizer.emis_vars[t].value()
    ccs_active = optimizer.active_vars[(t, 'CCS')].value()
    elec_cons = optimizer.cons_vars[(t, 'EN_ELEC')].value()
    print("Year", t, "Emis:", emis, "CCS:", ccs_active, "Elec:", elec_cons)
