# PathFinder Development Walkthrough

The PathFinder tool has been successfully developed and operationalized as a pipeline designed to minimize the Total Cost of Ownership (TCO) for industrial decarbonization while respecting constraints.

## 1. Modular Architecture Implemented

* **[model.py](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/model.py)**: Defines standard Python `@dataclasses` representing the internal state of the [PathFinderData](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/model.py#42-49) (Parameters, Resources, Technologies, Entities, Timelines).
* **[ingestion.py](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/ingestion.py)**: Contains the [PathFinderParser](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/ingestion.py#9-388) which uses `pandas` to read the complex, semi-structured `PathFinder input.xlsx`.
    * Discovers `START` and `END` blocks automatically.
    * Parses varying entities, baseline resource consumption, and linear interpolates time-series data using `np.interp`.
    * Re-aligns internal metrics: `kgCO2` to `tCO2` and standardizes energy to `MWh` equivalents. 
* **[optimizer.py](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/optimizer.py)**: Implements the [PathFinderOptimizer](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/optimizer.py#4-174) using the `PuLP` library.
    * Formulates the MILP problem.
    * Defines Decision Variables: binary investment flags ($Invest_{t, tech}$), temporal activity states ($Active_{t, tech}$), numerical consumptions ($Cons_{t, res}$), and Emissions ($Emis_{t, CO2}$).
    * Balances physical constraints representing the resource substitutions per unit.
* **[reporting.py](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/reporting.py)**: Exports the mathematically solved variables.
    * Integrates with `matplotlib` to generate [Energy_Mix.png](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/Energy_Mix.png) and [CO2_Trajectory.png](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/CO2_Trajectory.png).
    * Generates [Master_Plan.xlsx](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/Master_Plan.xlsx) with three granular sheets: Investments Schedule, Energy Mix tracking, and CO2 Trajectory factoring free quotas and carbon taxes.
* **[main.py](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/main.py)**: The orchestration file that seamlessly threads the components. 

## 2. Testing and Validation

* **Parser Robustness**: Tested on the `PathFinder input.xlsx` sheet structures successfully reading technologies (CCS, ERH, PEM_H2, FUEL_TO_H2, UP) without hardcoding absolute cell indices.
* **Model Feasibility**: Initially met with `Infeasible` statuses due to negative baseline resource parameters, this was debugged and relaxed to permit non-negative boundaries strictly for consumption/production variations. 
* **Solver Convergence**: The CBC Engine converges successfully, yielding optimal schedules (e.g. deciding `FUEL_TO_H2` investment scaling in 2025). 

## Output Execution

You can run the full optimization pipeline from your environment by executing:
```sh
.\venv\Scripts\python.exe main.py
```

The resulting files ([Master_Plan.xlsx](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/Master_Plan.xlsx), [Energy_Mix.png](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/Energy_Mix.png), and [CO2_Trajectory.png](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool/CO2_Trajectory.png)) are located directly inside `c:\Users\flori\Documents\TFM\PathWay_Python_Tool`.
