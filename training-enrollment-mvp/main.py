from __future__ import annotations

import importlib.util
from pathlib import Path


def _load_cli_main():
    """Load CLI entrypoint with a file-path fallback for Windows env quirks."""
    try:
        from app.cli import main as cli_main

        return cli_main
    except ModuleNotFoundError as exc:
        if exc.name not in {"app", "app.cli"}:
            raise

        base_dir = Path(__file__).resolve().parent
        cli_file = base_dir / "app" / "cli.py"
        if not cli_file.exists():
            raise ModuleNotFoundError(
                "Cannot find app.cli. Ensure you're in the project root and files are complete."
            ) from exc

        spec = importlib.util.spec_from_file_location("app_cli_fallback", cli_file)
        if spec is None or spec.loader is None:
            raise RuntimeError(f"Failed to load CLI module from {cli_file}") from exc

        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module.main


if __name__ == "__main__":
    _load_cli_main()()
