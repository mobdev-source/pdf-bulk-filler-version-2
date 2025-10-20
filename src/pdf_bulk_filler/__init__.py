"""Core package for pdf-bulk-filler."""

from .main import main, run_cli
from .ui.main_window import MainWindow

__all__ = ["main", "run_cli", "MainWindow"]
