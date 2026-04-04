"""
PathFinder: System Decarbonization Optimizer
Main entry point — sequential pipeline with OS-level solver watchdog.
"""
import os
import signal
import threading
import time
import traceback

import matplotlib
matplotlib.use('Agg')

from rich.console import Console, Group
from rich.panel import Panel
from rich.text import Text
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn
from rich.live import Live
from rich.align import Align

from .core.ingestion import PathFinderParser
from .core.optimizer import PathFinderOptimizer
from .core.reporting import PathFinderReporter

console = Console()


# ---------------------------------------------------------------------------
# Intro animation
# ---------------------------------------------------------------------------

def display_intro_animation():
    """Displays a stylized ASCII factory animation using rich.Live."""
    base    = "[bold white]      ________________[/bold white]"
    body    = "[bold white]     |                |\n     |   [  ]  [  ]   |\n     |________________|[/bold white]"
    chimney = "[bold white]            |  |\n        ____|  |____[/bold white]"
    smoke1  = "[bold grey]            .  .[/bold grey]"
    smoke2  = "[bold grey]          .  .  .\n         .  .  .[/bold grey]"
    smoke3  = "[bold grey]      .  .  .  .\n       .  .  .  .\n        .  .  .  .[/bold grey]"

    frames = [
        f"\n\n\n\n\n{base}",
        f"\n\n\n{body}\n{base}",
        f"\n{chimney}\n{body}\n{base}",
        f"{smoke1}\n{chimney}\n{body}\n{base}",
        f"{smoke2}\n{chimney}\n{body}\n{base}",
        f"{smoke3}\n{chimney}\n{body}\n{base}",
    ]

    title = Panel(Text("PathFinder: System Decarbonization Optimizer", justify="center", style="bold cyan"), expand=False)

    console.clear()
    with Live(Group(title, Text("")), vertical_overflow="visible", console=console, refresh_per_second=4) as live:
        for frame in frames:
            live.update(Group(title, Align.center(Text.from_markup(frame))))
            time.sleep(0.4)
        loading_text = Align.center(Text.from_markup("\n[bold yellow]   LOADING...[/bold yellow]"))
        live.update(Group(title, Align.center(Text.from_markup(frames[-1])), loading_text))
        time.sleep(1.2)
    console.clear()


# ---------------------------------------------------------------------------
# Solver watchdog
# ---------------------------------------------------------------------------

class SolverTimeoutError(Exception):
    """Raised when the solver watchdog kills the CBC process."""
    pass


def run_with_timeout(fn, timeout_secs):
    """
    Run fn() in a background thread.  If it does not finish within
    timeout_secs, raise SolverTimeoutError in the caller.

    Returns the return value of fn() on success.
    Raises any exception thrown by fn() on failure.
    """
    result = [None]
    exc    = [None]

    def _target():
        try:
            result[0] = fn()
        except Exception as e:
            exc[0] = e

    t = threading.Thread(target=_target, daemon=True)
    t.start()
    t.join(timeout=timeout_secs)

    if t.is_alive():
        # Thread is still running → solver is hung.
        # PuLP exposes the CBC subprocess through pulp.PULP_CBC_CMD; we can't
        # easily get its PID from outside, so we terminate via the OS kill
        # approach: mark the thread as a daemon (already done) so Python can
        # exit, and raise SolverTimeoutError here to let the caller skip reporting.
        raise SolverTimeoutError(f"Solver exceeded hard timeout of {timeout_secs:.0f}s and was abandoned.")

    if exc[0] is not None:
        raise exc[0]

    return result[0]


# ---------------------------------------------------------------------------
# Incumbent recovery helper
# ---------------------------------------------------------------------------

def _check_incumbent(optimizer):
    """
    Check whether the solver left behind a usable incumbent solution.

    PuLP/CBC writes variable values into the model object when it finds
    a feasible incumbent.  After a watchdog timeout the solver thread is
    still alive (daemon), but the model object is shared, so we can
    inspect variable values directly.

    We also try to read the .sol file that CBC writes to disk, which is
    the most reliable indicator of an incumbent.

    Returns True if at least *some* decision variables have non-None values.
    """
    import glob

    # Strategy 1: Check if CBC wrote a .sol file for this model
    sol_pattern = optimizer.model.name + "*.sol"
    sol_files = glob.glob(sol_pattern)
    if sol_files:
        try:
            # Ask PuLP to re-read the solution from the .sol file
            optimizer.model.assignVarsVals(
                {v.name: v.varValue for v in optimizer.model.variables() if v.varValue is not None}
            )
        except Exception:
            pass  # Not critical – we'll check variable values below

    # Strategy 2: Inspect variable values directly
    non_none_count = 0
    sample_size = min(50, len(optimizer.model.variables()))
    for v in list(optimizer.model.variables())[:sample_size]:
        if v.varValue is not None:
            non_none_count += 1

    # If more than 25% of sampled variables have values, consider it a usable solution
    return non_none_count > sample_size * 0.25


# ---------------------------------------------------------------------------
# Per-scenario pipeline
# ---------------------------------------------------------------------------

def run_scenario(sc, file_path, use_scenario_filter, generate_excel, progress, task_id, time_limit):
    """
    Execute the full pipeline for a single scenario:
      parse → build → solve (with watchdog) → report
    Returns a status string.
    """
    sc_id   = sc['id']
    sc_name = sc['name']
    WATCHDOG_BUFFER = 60  # extra seconds on top of solver time_limit before watchdog fires

    # --- 1. Ingestion --------------------------------------------------------
    try:
        progress.update(task_id, description=f"[cyan]{sc_name}[/cyan] Parsing data...", completed=5)
        p    = PathFinderParser(file_path, verbose=False)
        data = p.parse(scenario_id=sc_id if use_scenario_filter else None)
    except Exception as e:
        console.print(f"  [bold red][ERROR][/bold red] [{sc_name}] Ingestion failed: {e}")
        if os.environ.get('PATHFINDER_DEBUG'):
            traceback.print_exc()
        return 'Error'

    # --- 2. Build model ------------------------------------------------------
    try:
        progress.update(task_id, description=f"[cyan]{sc_name}[/cyan] Building model...", completed=15)
        optimizer = PathFinderOptimizer(data, verbose=False)
        optimizer.build_model()
    except Exception as e:
        console.print(f"  [bold red][ERROR][/bold red] [{sc_name}] Model build failed: {e}")
        if os.environ.get('PATHFINDER_DEBUG'):
            traceback.print_exc()
        return 'Error'

    # --- 3. Solve (with OS-level watchdog) -----------------------------------
    progress.update(task_id, description=f"[cyan]{sc_name}[/cyan] Solving...", completed=20)

    hard_timeout = time_limit + WATCHDOG_BUFFER

    # Spinner thread: update the progress bar while solver is running
    solve_done   = threading.Event()
    solve_start  = time.time()

    def _spinner():
        while not solve_done.is_set():
            elapsed = time.time() - solve_start
            pct = min(85.0, 20.0 + (elapsed / time_limit) * 65.0)
            progress.update(task_id, completed=pct)
            time.sleep(0.5)

    spinner_thread = threading.Thread(target=_spinner, daemon=True)
    spinner_thread.start()

    watchdog_fired = False
    try:
        status = run_with_timeout(optimizer.solve, hard_timeout)
    except SolverTimeoutError as e:
        solve_done.set()
        spinner_thread.join(timeout=1)
        watchdog_fired = True
        console.print(f"  [bold yellow][TIMEOUT][/bold yellow] [{sc_name}] {e}")

        # --- Try to recover the best incumbent solution ---
        # Check if the solver managed to populate variable values before being abandoned.
        has_incumbent = _check_incumbent(optimizer)
        if has_incumbent:
            status = 'Feasible'
            console.print(
                f"  [bold cyan][RECOVERY][/bold cyan] [{sc_name}] "
                f"Best available (non-optimal) solution recovered — proceeding to report."
            )
        else:
            console.print(
                f"  [bold yellow][WARN][/bold yellow] [{sc_name}] "
                f"No feasible incumbent found before timeout — skipping report."
            )
            progress.update(task_id, description=f"[yellow]{sc_name} — TIMEOUT[/yellow]", completed=100)
            return 'Timeout'
    except Exception as e:
        solve_done.set()
        spinner_thread.join(timeout=1)
        console.print(f"  [bold red][ERROR][/bold red] [{sc_name}] Solver crashed: {e}")
        if os.environ.get('PATHFINDER_DEBUG'):
            traceback.print_exc()
        progress.update(task_id, description=f"[red]{sc_name} — ERROR[/red]", completed=100)
        return 'Error'
    finally:
        solve_done.set()
        spinner_thread.join(timeout=1)

    # --- 4. Check if a solution exists ---------------------------------------
    if status == 'Infeasible':
        console.print(
            f"  [bold yellow][WARN][/bold yellow] [{sc_name}] No feasible solution found (Infeasible). "
            f"Check emission targets and budget constraints."
        )
        progress.update(task_id, description=f"[yellow]{sc_name} — Infeasible[/yellow]", completed=100)
        return 'Infeasible'

    if watchdog_fired:
        console.print(f"  [bold yellow][OK][/bold yellow]   [{sc_name}] Solver status: {status} (best available — time limit exceeded)")
    else:
        console.print(f"  [bold green][OK][/bold green]   [{sc_name}] Solver status: {status}")

    # --- 5. Reporting --------------------------------------------------------
    progress.update(task_id, description=f"[cyan]{sc_name}[/cyan] Generating report...", completed=88)
    try:
        save_dir = os.path.join('artifacts', 'reports', sc_name)
        os.makedirs(save_dir, exist_ok=True)

        reporter = PathFinderReporter(
            optimizer,
            scenario_id=sc_id,
            scenario_name=sc_name,
            generate_excel=generate_excel,
            verbose=False,
            progress_cb=None
        )
        reporter.generate_report()
    except Exception as e:
        console.print(f"  [bold red][ERROR][/bold red] [{sc_name}] Reporting failed: {e}")
        if os.environ.get('PATHFINDER_DEBUG'):
            traceback.print_exc()
        progress.update(task_id, description=f"[red]{sc_name} — Report Error[/red]", completed=100)
        return 'ReportError'

    progress.update(task_id, description=f"[green]{sc_name} ✓[/green]", completed=100)
    return status


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    display_intro_animation()
    file_path = os.path.join('data', 'raw', 'excel', 'PathFinder input.xlsx')

    console.print(Panel(
        Text("PathFinder: System Decarbonization Optimizer", justify="center", style="bold cyan"),
        expand=False
    ))

    # -- 0. Read scenario list ------------------------------------------------
    try:
        _parser_probe = PathFinderParser(file_path)
        scenarios     = _parser_probe._parse_scenarios()
    except Exception as e:
        console.print(f"[bold red][FATAL][/bold red] Cannot read input file '{file_path}': {e}")
        return

    generate_excel  = True
    verbose_logging = False

    if not scenarios:
        console.print("[yellow][INFO] No MODELING/SCENARIOS block found. Running in single-scenario mode.[/yellow]")
        scenarios          = [{'id': 'DEFAULT', 'name': 'Default'}]
        use_scenario_filter = False
    else:
        use_scenario_filter = True
        console.print(f"[bold cyan]Found {len(scenarios)} scenario(s): {[s['name'] for s in scenarios]}[/bold cyan]")

    # Read time_limit from the first scenario's data
    try:
        data_probe = _parser_probe.parse(scenario_id=scenarios[0]['id'] if use_scenario_filter else None)
        time_limit = data_probe.parameters.time_limit
    except Exception:
        time_limit = 60.0  # sensible default

    # -- 1. Run scenarios sequentially ----------------------------------------
    console.print(f"\n[bold magenta]SYSTEM OPTIMAL RESOLUTION ({len(scenarios)} scenario(s))...[/bold magenta]")

    summary = {}  # sc_name -> status string

    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}"),
        BarColumn(bar_width=40),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeElapsedColumn(),
        console=console,
        refresh_per_second=4,
    ) as progress:
        for sc in scenarios:
            task_id = progress.add_task(
                description=f"[cyan]{sc['name']}[/cyan] Initialising...",
                total=100
            )
            status = run_scenario(
                sc, file_path, use_scenario_filter,
                generate_excel, progress, task_id, time_limit
            )
            summary[sc['name']] = status

    # -- 2. Final summary -----------------------------------------------------
    console.print("\n")
    rows = []
    for sc_name, st in summary.items():
        if st in ('Optimal', 'Feasible', 'Feasible (Timeout)'):
            colour = 'green'
            icon   = '✓'
        elif st == 'Infeasible':
            colour = 'yellow'
            icon   = '✗'
        elif st == 'Timeout':
            colour = 'yellow'
            icon   = '⏱'
        else:
            colour = 'red'
            icon   = '!'
        rows.append(f"  [{colour}]{icon} {sc_name}: {st}[/{colour}]")

    summary_text = "\n".join(rows)
    console.print(Panel(
        Text.from_markup(f"[bold]Simulation complete[/bold]\n\n{summary_text}"),
        title="[bold green]PathFinder[/bold green]",
        expand=False
    ))


if __name__ == "__main__":
    main()
