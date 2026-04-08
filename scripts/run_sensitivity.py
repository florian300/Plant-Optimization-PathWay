"""
run_sensitivity.py — CLI Wrapper for OAT Sensitivity Analysis
============================================================
This script is now a thin wrapper around the core sensitivity engine.
Usage:
    python scripts/run_sensitivity.py --excel-path "PathFinder input.xlsx"
"""

import argparse
import os
import sys
from pathlib import Path

# --- Source Path Bootstrapping ---
_REPO_ROOT = Path(__file__).resolve().parents[1]
_SRC_DIR   = _REPO_ROOT / "src"
if str(_SRC_DIR) not in sys.path:
    sys.path.insert(0, str(_SRC_DIR))

from pathway.core.sensitivity_engine import run_sensitivity

def main():
    parser = argparse.ArgumentParser(description="OAT Sensitivity Analysis — PathFinder")
    parser.add_argument("--excel-path", type=str, default=str(_REPO_ROOT / "data" / "raw" / "excel" / "PathFinder input.xlsx"))
    parser.add_argument("--output", type=str, default=None)
    parser.add_argument("--verbose", "-v", action="store_true")
    args = parser.parse_args()

    output_path = args.output
    if output_path is None:
        output_path = str(_REPO_ROOT / "artifacts" / "sensitivity" / "sensitivity_results.json")

    run_sensitivity(
        excel_path=args.excel_path,
        output_path=output_path,
        verbose=args.verbose
    )

if __name__ == "__main__":
    main()
