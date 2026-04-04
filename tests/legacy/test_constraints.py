from main import *
from ingestion import PathFinderParser
from optimizer import PathFinderOptimizer
import pulp

parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()

opt = PathFinderOptimizer(data)
opt.build_model()
opt.solve()
print("Final Status:", pulp.LpStatus[opt.model.status])
