"""
Solver Factory for PathFinder Optimization Engine.

Provides a decoupled, configuration-driven solver instantiation mechanism
supporting HiGHS, Gurobi, CPLEX, and CBC with graceful fallback.
"""

import logging
from typing import Optional

import pulp

logger = logging.getLogger(__name__)

# Ordered preference for automatic fallback resolution.
_SOLVER_REGISTRY = {
    "HIGHS": ["HiGHS", "HiGHS_CMD"],  # Try library first, then CMD
    "GUROBI": ["GUROBI_CMD", "GUROBI"],
    "CPLEX": ["CPLEX_CMD", "CPLEX"],
    "CBC": ["PULP_CBC_CMD"],
}


def get_solver(
    solver_name: str,
    time_limit: int,
    gap_rel: float,
    msg: bool = True,
    threads: int = 4,
    warm_start: bool = False,
) -> pulp.LpSolver:
    """Instantiate a PuLP solver by name with graceful fallback to CBC.

    Parameters
    ----------
    solver_name : str
        Desired solver identifier. One of ``'HIGHS'``, ``'GUROBI'``,
        ``'CPLEX'``, or ``'CBC'`` (case-insensitive).
    time_limit : int
        Maximum wall-clock seconds allowed for the solve.
    gap_rel : float
        Relative MIP gap tolerance (e.g. ``0.01`` for 1 %).
    msg : bool, optional
        Whether the solver should print its native log output.
        Defaults to ``True``.
    threads : int, optional
        Number of parallel threads to use. Defaults to ``4``.
    warm_start : bool, optional
        Whether to enable MIP warm-start. Defaults to ``False``.

    Returns
    -------
    pulp.LpSolver
        A ready-to-use solver instance.

    Notes
    -----
    If the requested solver is not installed or cannot be located, the
    function logs a warning and silently falls back to the PuLP-bundled
    CBC solver, guaranteeing that solves never fail due to a missing
    binary.
    """

    key = solver_name.strip().upper()
    solver: Optional[pulp.LpSolver] = None

    if key not in _SOLVER_REGISTRY:
        logger.warning(
            "Unknown solver '%s'. Recognised values: %s. Falling back to CBC.",
            solver_name,
            list(_SOLVER_REGISTRY.keys()),
        )
        key = "CBC"

    # --- Attempt to build the requested solver --------------------------------
    if key != "CBC":
        solver = _try_build_solver(
            key,
            time_limit=time_limit,
            gap_rel=gap_rel,
            msg=msg,
            threads=threads,
            warm_start=warm_start,
        )

    # --- Fallback to CBC if the primary choice failed -------------------------
    if solver is None:
        if key != "CBC":
            logger.warning(
                "Solver '%s' is not available. Falling back to CBC.", key
            )
        solver = _build_cbc(
            time_limit=time_limit,
            gap_rel=gap_rel,
            msg=msg,
            threads=threads,
            warm_start=warm_start,
        )

    return solver


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _try_build_solver(
    key: str,
    *,
    time_limit: int,
    gap_rel: float,
    msg: bool,
    threads: int,
    warm_start: bool,
) -> Optional[pulp.LpSolver]:
    """Attempt to build a non-CBC solver. Returns ``None`` on failure."""

    pulp_names = _SOLVER_REGISTRY[key]
    if isinstance(pulp_names, str):
        pulp_names = [pulp_names]

    for pulp_name in pulp_names:
        try:
            solver = pulp.getSolver(
                pulp_name,
                msg=msg,
                timeLimit=time_limit,
                gapRel=gap_rel,
                threads=threads,
                warmStart=warm_start,
            )

            if not solver.available():
                logger.debug(
                    "Solver class '%s' instantiated but reports unavailable.",
                    pulp_name,
                )
                continue

            logger.info("Solver '%s' (%s) selected successfully.", key, pulp_name)
            return solver

        except Exception as exc:  # noqa: BLE001
            logger.debug(
                "Could not instantiate solver '%s': %s", pulp_name, exc
            )
            continue
            
    return None


def _build_cbc(
    *,
    time_limit: int,
    gap_rel: float,
    msg: bool,
    threads: int,
    warm_start: bool,
) -> pulp.LpSolver:
    """Build the bundled PuLP CBC solver (always available)."""

    solver = pulp.PULP_CBC_CMD(
        msg=msg,
        timeLimit=time_limit,
        gapRel=gap_rel,
        threads=threads,
        presolve=True,
        warmStart=warm_start,
        keepFiles=warm_start,  # Required for warmStart on Windows
    )
    logger.info("CBC solver instantiated (time_limit=%s, gap_rel=%s).", time_limit, gap_rel)
    return solver
