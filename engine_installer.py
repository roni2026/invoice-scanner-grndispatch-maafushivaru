"""
engine_installer.py
===================
OCR Engine Installer Utility — Maafushivaru Document Hub
Checks whether OCR engines are installed and installs them
on-demand via pip. Used by both the GUI and CLI tools.

Author : Roni  |  Version: 3.5
"""

import sys
import subprocess
import importlib
from typing import Callable, Optional

# ─────────────────────────────────────────────────────────────
# ENGINE REGISTRY
# ─────────────────────────────────────────────────────────────
ENGINE_LABELS = {
    "tesseract": "Tesseract OCR",
    "paddleocr": "PaddleOCR",
    "easyocr":   "EasyOCR",
}

# pip packages required to run each engine
ENGINE_PACKAGES = {
    "tesseract": [],                              # Binary install — no pip
    "paddleocr": ["paddlepaddle", "paddleocr"],
    "easyocr":   ["easyocr"],
}

# Python module to import-check for each engine
ENGINE_MODULES = {
    "tesseract": "pytesseract",
    "paddleocr": "paddleocr",
    "easyocr":   "easyocr",
}

# Approximate download sizes shown to the user before install
ENGINE_SIZES = {
    "tesseract": "Binary installer — see link below",
    "paddleocr": "~300 MB  (model files download on first use)",
    "easyocr":   "~100 MB  (model files download on first use)",
}

# Download links for engines that need a binary installer
ENGINE_LINKS = {
    "tesseract": "https://github.com/UB-Mannheim/tesseract/wiki",
}


# ─────────────────────────────────────────────────────────────
# CHECK
# ─────────────────────────────────────────────────────────────
def is_installed(engine: str) -> bool:
    """
    Return True if the Python package for the engine can be imported.
    Uses importlib so the check always reflects the current environment.
    """
    module = ENGINE_MODULES.get(engine, engine)
    try:
        importlib.import_module(module)
        return True
    except ImportError:
        return False


def get_packages(engine: str) -> list:
    """Return the list of pip packages needed for this engine."""
    return ENGINE_PACKAGES.get(engine, [])


def get_install_command(engine: str) -> str:
    """Return the pip install command string for display purposes."""
    pkgs = get_packages(engine)
    if not pkgs:
        return f"See: {ENGINE_LINKS.get(engine, 'N/A')}"
    return f"pip install {' '.join(pkgs)}"


# ─────────────────────────────────────────────────────────────
# INSTALL
# ─────────────────────────────────────────────────────────────
def install_engine(
    engine: str,
    line_callback: Optional[Callable[[str], None]] = None
) -> tuple:
    """
    Install the pip packages for the specified OCR engine.

    Args:
        engine        : Engine key  ('paddleocr'  or  'easyocr')
        line_callback : Optional callable(line: str) called for each output
                        line from pip. Use this to stream progress to a UI.

    Returns:
        (success: bool, message: str)
    """
    if engine == "tesseract":
        return False, (
            "Tesseract is a system binary — it cannot be installed via pip.\n"
            f"Download from:  {ENGINE_LINKS['tesseract']}")

    packages = get_packages(engine)
    if not packages:
        return True, "No pip installation needed for this engine."

    cmd = [sys.executable, "-m", "pip", "install", "--upgrade"] + packages

    # On Windows, suppress the console window when called from a GUI app
    creation_flags = 0
    if sys.platform == "win32":
        creation_flags = subprocess.CREATE_NO_WINDOW

    try:
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            creationflags=creation_flags,
        )

        for line in process.stdout:
            stripped = line.rstrip()
            if stripped and line_callback:
                line_callback(stripped)

        process.wait()

        # Refresh importlib cache so newly installed packages are discoverable
        importlib.invalidate_caches()

        if process.returncode == 0:
            return True, f"Successfully installed: {', '.join(packages)}"
        else:
            return False, (
                f"pip exited with code {process.returncode}.\n"
                f"Try running manually:  {get_install_command(engine)}")

    except FileNotFoundError:
        return False, (
            "Python executable not found in PATH.\n"
            f"Try manually:  {get_install_command(engine)}")
    except Exception as e:
        return False, f"Unexpected error during installation: {e}"


# ─────────────────────────────────────────────────────────────
# CLI HELPER  (used by dashboard.py and grndispatch.py)
# ─────────────────────────────────────────────────────────────
def cli_install_engine(engine: str, console) -> bool:
    """
    Interactive CLI install flow using Rich console output.
    Streams pip output line-by-line with live display.

    Args:
        engine  : Engine key
        console : rich.console.Console instance

    Returns:
        True if installation succeeded, False otherwise.
    """
    from rich.live  import Live
    from rich.panel import Panel
    from rich.text  import Text

    label   = ENGINE_LABELS.get(engine, engine)
    pkgs    = get_packages(engine)
    cmd_str = get_install_command(engine)
    size    = ENGINE_SIZES.get(engine, "")

    console.print(f"\n  [bold cyan]Installing {label}[/bold cyan]")
    console.print(f"  Command  : [dim]{cmd_str}[/dim]")
    console.print(f"  Size     : [dim]{size}[/dim]")
    console.print(f"  [dim]Streaming pip output below...[/dim]\n")

    output_lines: list = []

    def on_line(line: str):
        output_lines.append(line)
        # Only show the last line live to keep it clean
        console.print(f"  [dim]{line}[/dim]")

    success, message = install_engine(engine, line_callback=on_line)

    console.print()
    if success:
        console.print(f"  [bold green]✔  {message}[/bold green]")
    else:
        console.print(f"  [bold red]✖  {message}[/bold red]")

    return success