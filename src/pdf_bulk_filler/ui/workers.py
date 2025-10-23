"""Background workers used by the Qt interface."""

from __future__ import annotations

from pathlib import Path
import tempfile
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional

from PySide6 import QtCore

import fitz
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import BooleanObject, DictionaryObject, NameObject, NumberObject, TextStringObject

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
        rows: Iterable[Mapping[str, Any]],
        *,
        flatten: bool = False,
        read_only: bool = False,
        template_metadata: PdfTemplate | None = None,
        mode: str = "per_entry",
        combined_output: Path | None = None,
        filename_builder: Optional[Callable[[Mapping[str, Any], int], str]] = None,
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
        self._mode = "combined" if mode == "combined" else "per_entry"
        self._combined_output = combined_output
        self._filename_builder = filename_builder

    @QtCore.Slot()
    def run(self) -> None:
        try:
            if self._mode == "combined":
                outputs = self._generate_combined()
            else:
                outputs = self._generate_individual()
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

    def _generate_individual(self) -> List[Path]:
        destination = self._output_dir
        destination.mkdir(parents=True, exist_ok=True)
        outputs = self._engine.fill_rows(
            self._template_path,
            destination,
            self._rules,
            self._rows,
            filename_builder=self._filename_builder,
            progress_callback=self._report_progress,
            flatten=self._flatten,
            template_metadata=self._template_metadata,
            read_only=self._read_only,
        )
        for path in outputs:
            self._refresh_widget_appearances(path)
        return outputs

    def _generate_combined(self) -> List[Path]:
        if self._combined_output is None:
            raise ValueError("Combined output path was not provided.")
        final_path = self._combined_output.with_suffix(".pdf")
        final_path.parent.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory() as tmpdir:
            temp_dir = Path(tmpdir)
            outputs = self._engine.fill_rows(
                self._template_path,
                temp_dir,
                self._rules,
                self._rows,
                filename_builder=self._filename_builder,
                progress_callback=self._report_progress,
                flatten=self._flatten,
                template_metadata=self._template_metadata,
                read_only=self._read_only,
            )

            combined_writer = PdfWriter()
            for index, pdf_path in enumerate(outputs, start=1):
                suffix = f"entry{index:04d}"
                with pdf_path.open("rb") as input_handle:
                    reader = PdfReader(input_handle)
                    self._rename_form_fields(reader, suffix)
                    if self._read_only:
                        self._set_read_only(reader)
                    combined_writer.append_pages_from_reader(reader)

        acro_ref = combined_writer._root_object.get(NameObject("/AcroForm"))  # type: ignore[attr-defined]
        if acro_ref is not None:
            acro_form = acro_ref.get_object() if hasattr(acro_ref, "get_object") else acro_ref
            acro_form.update({NameObject("/NeedAppearances"): BooleanObject(True)})

        with final_path.open("wb") as output_handle:
            combined_writer.write(output_handle)

        self._refresh_widget_appearances(final_path)
        return [final_path]

    def _set_read_only(self, reader: PdfReader) -> None:
        for page in reader.pages:
            annotations = page.get("/Annots")
            if not annotations:
                continue
            annots = annotations.get_object() if hasattr(annotations, "get_object") else annotations
            for annot_ref in list(annots):
                annot = annot_ref.get_object()
                flags = int(annot.get("/Ff", 0))
                annot[NameObject("/Ff")] = NumberObject(flags | 1)

    def _refresh_widget_appearances(self, pdf_path: Path) -> None:
        document = fitz.open(pdf_path)
        try:
            for page_index in range(document.page_count):
                page = document.load_page(page_index)
                for widget in page.widgets() or []:
                    try:
                        widget.update()
                    except Exception:  # noqa: BLE001
                        continue
                page.clean_contents()
            temp_path = pdf_path.with_suffix(".tmp.pdf")
            document.save(temp_path, garbage=3, deflate=True)
        finally:
            document.close()

        temp_path.replace(pdf_path)

    def _rename_form_fields(self, reader: PdfReader, suffix: str) -> None:
        field_map: Dict[str, str] = {}

        def _rename_field(field_obj: Any) -> None:
            if field_obj is None:
                return
            if hasattr(field_obj, "get_object"):
                field_obj = field_obj.get_object()
            if not isinstance(field_obj, dict):
                return

            original = field_obj.get("/T")
            if original:
                original_text = str(original)
                new_name = f"{original_text}_{suffix}"
                field_obj[NameObject("/T")] = TextStringObject(new_name)
                field_map[original_text] = new_name

            kids = field_obj.get("/Kids")
            if kids:
                for kid in kids:
                    _rename_field(kid)

        acro_ref = reader.trailer["/Root"].get("/AcroForm")
        if acro_ref is not None:
            acro_form = acro_ref.get_object() if hasattr(acro_ref, "get_object") else acro_ref
            fields = acro_form.get("/Fields")
            if fields:
                for field in list(fields):
                    _rename_field(field)
            acro_form.update({NameObject("/NeedAppearances"): BooleanObject(True)})

        for page in reader.pages:
            annotations = page.get("/Annots")
            if not annotations:
                continue
            annots = annotations.get_object() if hasattr(annotations, "get_object") else annotations
            for annot_ref in list(annots):
                annot = annot_ref.get_object()
                field_name = annot.get("/T")
                if not field_name:
                    continue

                name_str = str(field_name)
                new_name = field_map.get(name_str)
                if not new_name:
                    new_name = f"{name_str}_{suffix}"
                    field_map[name_str] = new_name
                annot[NameObject("/T")] = TextStringObject(new_name)

                if self._read_only:
                    flags = int(annot.get("/Ff", 0))
                    annot[NameObject("/Ff")] = NumberObject(flags | 1)
