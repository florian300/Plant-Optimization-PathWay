import pulp
import highspy
print(f"Version PuLP: {pulp.VERSION}")
print(f"Solveurs disponibles pour PuLP: {pulp.listSolvers(onlyAvailable=True)}")

# Test spécifique via votre Factory si possible
try:
    from src.pathway.core.solver_factory import get_solver
    solver = get_solver("HIGHS", time_limit=10, gap_rel=0.01, msg=False)
    print(f"Succès ! Solveur récupéré via la factory : {solver}")
except Exception as e:
    print(f"Erreur lors du test de la factory : {e}")
