"""Background workers used by the Qt interface."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, List

from PySide6 import QtCore

from pdf_bulk_filler.pdf.engine import PdfEngine, PdfTemplate
from pdf_bulk_filler.mapping.rules import coerce_rules


class PdfGenerationWorker(QtCore.QObject):
    """Run PDF generation in a background thread."""

    progress = QtCore.Signal(int, int)
    completed = QtCore.Signal(list)
    failed = QtCore.Signal(str)
    cancelled = QtCore.Signal()

    def __init__(
        self,
        engine: PdfEngine,
        template_path: Path,
        output_dir: Path,
        rule_spec: object,
        rows: Iterable[Dict[str, object]],
        *,
        flatten: bool = False,
        read_only: bool = False,
        template_metadata: PdfTemplate | None = None,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._engine = engine
        self._template_path = template_path
        self._output_dir = output_dir
        self._rules = coerce_rules(rule_spec)
        self._rows = list(rows)
        self._flatten = flatten
        self._read_only = read_only
        self._template_metadata = template_metadata
        self._cancel_requested = False

    @QtCore.Slot()
    def run(self) -> None:
        try:
            outputs = self._engine.fill_rows(
                self._template_path,
                self._output_dir,
                self._rules,
                self._rows,
                filename_pattern="{index:05d}",
                progress_callback=self._report_progress,
                flatten=self._flatten,
                template_metadata=self._template_metadata,
                read_only=self._read_only,
            )
        except KeyboardInterrupt:
            self.cancelled.emit()
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))
        else:
            self.completed.emit(outputs)

    def request_cancel(self) -> None:
        """Signal that the worker should abort as soon as possible."""
        self._cancel_requested = True

    def _report_progress(self, current: int, _: int) -> None:
        if self._cancel_requested:
            raise KeyboardInterrupt
        self.progress.emit(current, len(self._rows))
