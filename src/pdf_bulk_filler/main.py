"""Application entry point for the PDF bulk filler MVP."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys

from PySide6 import QtCore, QtWidgets

from pdf_bulk_filler.ui.main_window import MainWindow

_PROJECT_VERSION = "0.1.0"


def run_cli(args: list[str] | None = None) -> int:
    """Parse CLI arguments and dispatch commands."""
    parser = argparse.ArgumentParser(
        prog="pdf-bulk-filler",
        description="Desktop application for mapping tabular data onto PDF form templates.",
    )
    parser.add_argument("--version", action="store_true", help="Print the package version and exit.")
    parser.add_argument(
        "--no-gui",
        action="store_true",
        help="Skip launching the desktop interface (useful for CI smoke tests).",
    )
    parser.add_argument(
        "--template",
        type=Path,
        help="Optional PDF template to pre-load when the GUI starts.",
    )
    parser.add_argument(
        "--data",
        type=Path,
        help="Optional CSV/Excel source to pre-load when the GUI starts.",
    )
    parser.add_argument(
        "--mapping",
        type=Path,
        help="Optional JSON mapping configuration to load at startup.",
    )

    parsed = parser.parse_args(args=args)

    if parsed.version:
        print(f"pdf-bulk-filler {_PROJECT_VERSION}")
        return 0

    if parsed.no_gui:
        print("GUI launch suppressed via --no-gui.")
        return 0

    return launch_app(parsed)


def main() -> int:
    """Entrypoint used by console scripts."""
    return run_cli()


def launch_app(options: argparse.Namespace) -> int:
    """Create and run the Qt application."""
    app = QtWidgets.QApplication.instance()
    if app is None:
        app = QtWidgets.QApplication(sys.argv)

    window = MainWindow()
    window.show()

    # Defer post-init loading until the event loop starts to ensure widgets exist.
    def _post_init() -> None:
        if getattr(options, "mapping", None):
            window._action_load_mapping_from_path(options.mapping)  # type: ignore[attr-defined]
        else:
            if getattr(options, "data", None):
                window._action_import_data_from_path(Path(options.data))  # type: ignore[attr-defined]
            if getattr(options, "template", None):
                window._action_import_pdf_from_path(Path(options.template))  # type: ignore[attr-defined]

    QtCore.QTimer.singleShot(0, _post_init)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
