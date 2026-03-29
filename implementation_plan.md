# Multi-Scenario Support

The Excel file defines 3 scenarios (BS, CT, LCB) in the OverView `MODELING` block. Currently the tool only runs once. This change makes it run once per scenario, filtering all scenario-specific data accordingly and saving results in dedicated sub-folders.

## Proposed Changes

### Core — ingestion.py

#### [MODIFY] [ingestion.py](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool%20-%20v2/core/ingestion.py)

- Add `_parse_scenarios()` method: reads MODELING START/END block from OverView and returns `list[dict]` with `{id, name}` per scenario
- Add `scenario_id: str` parameter to [parse()](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool%20-%20v2/core/ingestion.py#112-1082) — when provided, filters all scenario-specific sheets
- **RESSOURCES_PRICE**: scan for `SCENARIO N START/END` blocks, read `SC-DES` inside each to match the scenario_id, then only parse that block's price data
- **OTHER EMISSIONS**: same logic as RESSOURCES_PRICE (already has same structure)
- **NEW TECH costs (COMPATIBILITIES block)**: filter rows where the `SCENARIO` column matches `scenario_id` (or is empty = common to all)
- **NEW TECH_INDIRECT CARAC rows**: filter rows where the scenario column matches `scenario_id`
- **REFINERY BUDGET**: find the `BUDGET <scenario_id>` row; fall back to a default if not found

---

### Core — main.py

#### [MODIFY] [main.py](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool%20-%20v2/main.py)

- After opening the file, call `PathFinderParser._parse_scenarios()` to get the scenario list
- Loop over each scenario, creating a [PathFinderParser](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool%20-%20v2/core/ingestion.py#12-1082) and calling [parse(scenario_id=...)](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool%20-%20v2/core/ingestion.py#112-1082) each time
- For each scenario: build model → solve → if Optimal, call [PathFinderReporter](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool%20-%20v2/core/reporting.py#11-2112) with `scenario_id` and `scenario_name`
- Console output identifies which scenario is running

---

### Core — reporting.py

#### [MODIFY] [reporting.py](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool%20-%20v2/core/reporting.py)

- Accept `scenario_id: str` and `scenario_name: str` in [__init__](file:///c:/Users/flori/Documents/TFM/PathWay_Python_Tool%20-%20v2/core/reporting.py#12-17)
- All `os.makedirs` and `plt.savefig` paths changed from `Results/` → `Results/<scenario_id>/`
- All chart filenames prefixed with `<scenario_id>_` (e.g. `BS_Energy_Mix_Consumption_MWh.png`)
- Excel output (`Master_Plan.xlsx`) also saved to `Results/<scenario_id>/`
- Add `_add_scenario_label()` helper: places a small text box in the top-right corner of every figure with the scenario name. Style: **white bold text**, colored background (unique per scenario ID using a fixed color palette)
- Call `_add_scenario_label()` before every `plt.savefig()` call

**Scenario color palette** (background of the label box):
| Scenario ID | Color |
|---|---|
| BS | `#1A5276` (dark blue) |
| CT | `#1E8449` (dark green) |
| LCB | `#6E2F7C` (purple) |
| Others | auto-cycle from palette |

---

## Verification Plan

### Automated Test
Run the tool and verify output structure:
```
cd "c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2"
python main.py
```

Then check:
```
python -c "
import os
for sc in ['BS', 'CT', 'LCB']:
    d = f'Results/{sc}'
    files = os.listdir(d) if os.path.exists(d) else []
    charts = [f for f in files if f.endswith('.png')]
    print(f'{sc}: {len(charts)} charts, prefix check: {all(f.startswith(sc) for f in charts)}')
"
```

Expected: 3 scenario folders, each with charts prefixed by scenario ID.

### Manual Verification
Open any output PNG from `Results/BS/`, `Results/CT/`, `Results/LCB/` and verify:
1. The chart filename starts with the scenario ID
2. A colored box in the **top-right** displays the scenario name in white bold text
3. Each scenario folder has its own `Master_Plan.xlsx`
