"""Compatibility entrypoint for running the app from repository root."""

from pathlib import Path
import sys


def _bootstrap_src_path() -> None:
    repo_root = Path(__file__).resolve().parent
    src_dir = repo_root / "src"
    src_path = str(src_dir)
    if src_path not in sys.path:
        sys.path.insert(0, src_path)


def main() -> None:
    _bootstrap_src_path()
    from pathway.cli import main as cli_main

    cli_main()


if __name__ == "__main__":
    main()
