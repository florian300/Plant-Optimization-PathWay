import sys
import os

# Add src to path
sys.path.append(os.path.abspath("src"))

from pathway.core.optimizer import PathFinderOptimizer
from pathway.core.model import PathFinderData, Resource, Parameters, EntityState

# Mock Data
class MockData:
    def __init__(self):
        self.resources = {
            "R1": Resource(id="R1", name="GJ_Resource", unit="GJ", type="ENERGY", resource_type="ENERGY"),
            "R2": Resource(id="R2", name="MWH_Resource", unit="MWH", type="ENERGY", resource_type="ENERGY"),
            "R3": Resource(id="R3", name="BBL_Resource", unit="BBL", type="ENERGY", resource_type="ENERGY"),
            "R4": Resource(id="R4", name="CO2_Resource", unit="TCO2", type="EMISSION", resource_type="CO2")
        }
        self.unit_conversions = {
             ("GJ", "MWH"): 1/3.6,
             ("MWH", "BBL"): 1/1.7
        }
        self.entities = {
            "E1": EntityState(id="E1", name="Test", base_consumptions={"R1": 3600}, processes={})
        }
        self.parameters = Parameters(start_year=2025, duration=1, entities=self.entities, resources=list(self.resources.keys()))
        self.technologies = {}
        self.time_series = None
        self.objectives = []
        self.grant_params = None
        self.ccfd_params = None
        self.bank_loans = []
        self.dac_params = None
        self.credit_params = None
        self.reporting_toggles = None

def test_unit_conversion():
    data = MockData()
    opt = PathFinderOptimizer(data)
    
    print("Testing Direct Lookup (GJ -> MWH)...")
    factor = opt._get_unit_conversion("R1", "R2")
    print(f"Result: {factor} (Expected: {1/3.6})")
    assert abs(factor - 1/3.6) < 1e-6
    
    print("Testing Reverse Lookup (BBL -> MWH)...")
    factor = opt._get_unit_conversion("R3", "R2")
    print(f"Result: {factor} (Expected: 1.7)")
    assert abs(factor - 1.7) < 1e-6
    
    print("Testing Cross-Resource Link (CO2 -> MWH)...")
    # TCO2 and MWH have different resource_types (CO2 and ENERGY)
    # Should return 1.0 (linkage)
    factor = opt._get_unit_conversion("R4", "R2")
    print(f"Result: {factor} (Expected: 1.0)")
    assert factor == 1.0

    print("Testing ValueError for intra-domain mismatch (GJ -> BBL)...")
    try:
        opt._get_unit_conversion("R1", "R3")
        print("FAILED: Should have raised ValueError")
    except ValueError as e:
        print(f"SUCCESS: Caught expected error: {e}")

if __name__ == "__main__":
    try:
        test_unit_conversion()
        print("\nALL UNIT FIX TESTS PASSED!")
    except Exception as e:
        print(f"\nTESTS FAILED: {e}")
        sys.exit(1)
