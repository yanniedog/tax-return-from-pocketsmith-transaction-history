#!/usr/bin/env python3
"""One-command launcher for the PocketSmith tax prep app."""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent


def command_exists(name: str) -> bool:
    return shutil.which(name) is not None


def run_checked(command: list[str]) -> None:
    subprocess.run(command, cwd=PROJECT_ROOT, check=True)


def _npm_cmd(*args: str) -> list[str]:
    """Return [npm_path, ...args]. Use after ensure_node_tooling() so npm is in PATH."""
    npm_path = shutil.which("npm")
    if not npm_path:
        raise FileNotFoundError("npm not found in PATH")
    return [npm_path, *args]


def ensure_node_tooling() -> None:
    missing = [name for name in ("node", "npm") if not command_exists(name)]
    if not missing:
        return
    missing_text = ", ".join(missing)
    print(f"Missing required command(s): {missing_text}", file=sys.stderr)
    print("Install Node.js from https://nodejs.org/ and retry.", file=sys.stderr)
    sys.exit(1)


def ensure_dependencies() -> None:
    node_modules = PROJECT_ROOT / "node_modules"
    if node_modules.exists():
        return
    print("Installing dependencies...")
    run_checked(_npm_cmd("install"))


def start_app() -> int:
    print("Starting PocketSmith Tax Prep...")
    process = subprocess.Popen(_npm_cmd("start"), cwd=PROJECT_ROOT)
    try:
        return process.wait()
    except KeyboardInterrupt:
        print("\nStopping app...")
        process.terminate()
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
        return 130


def main() -> int:
    ensure_node_tooling()
    ensure_dependencies()
    return start_app()


if __name__ == "__main__":
    raise SystemExit(main())
