from core.ingestion import PathFinderParser
from core.optimizer import PathFinderOptimizer
import pulp
import time

file_path = 'PathFinder input.xlsx'
p = PathFinderParser(file_path, verbose=False)
# Run for CT as it was the slow one
data = p.parse(scenario_id='CT')

start = time.time()
print("Building model...")
optimizer = PathFinderOptimizer(data, verbose=False)
optimizer.build_model()
build_time = time.time() - start
print(f"Model built in {build_time:.2f}s")

# Check if relaxation is working
# Look at some invest_vars for a process with nb_units > 1
relaxed_count = 0
total_count = 0
for v in optimizer.invest_vars.values():
    total_count += 1
    if v.cat == pulp.LpContinuous:
        relaxed_count += 1

print(f"Relaxed variables: {relaxed_count}/{total_count}")

print("Solving...")
start = time.time()
status = optimizer.solve()
solve_time = time.time() - start
print(f"Solved in {solve_time:.2f}s with status {status}")
