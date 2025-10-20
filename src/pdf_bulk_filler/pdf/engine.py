"""PDF utilities for loading templates and generating filled documents."""

from __future__ import annotations

from dataclasses import dataclass
import io
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import fitz  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject


@dataclass(frozen=True)
class PdfField:
    """Description of a PDF form field."""

    field_name: str
    page_index: int
    rect: fitz.Rect


@dataclass
class PdfTemplate:
    """Container for an opened PDF template."""

    path: Path
    document: fitz.Document
    fields: List[PdfField]

    def close(self) -> None:
        """Close the underlying document."""
        self.document.close()


class PdfEngine:
    """High-level operations for reading, rendering, and filling PDF templates."""

    def open_template(self, path: Path) -> PdfTemplate:
        template_path = path.expanduser().resolve()
        if not template_path.exists():
            raise FileNotFoundError(f"PDF template not found: {template_path}")

        document = fitz.open(template_path)
        fields: List[PdfField] = []
        for page_index in range(document.page_count):
            page = document.load_page(page_index)
            for widget in page.widgets() or []:
                if widget.field_name:
                    fields.append(
                        PdfField(
                            field_name=widget.field_name,
                            page_index=page_index,
                            rect=widget.rect,
                        )
                    )
        return PdfTemplate(path=template_path, document=document, fields=fields)

    def render_page(self, template: PdfTemplate, page_index: int, zoom: float = 1.5) -> fitz.Pixmap:
        """Render a page to a pixmap for display in the UI."""
        page = template.document.load_page(page_index)
        matrix = fitz.Matrix(zoom, zoom)
        return page.get_pixmap(matrix=matrix, alpha=False)

    def fill_rows(
        self,
        template_path: Path,
        destination_dir: Path,
        field_mapping: Dict[str, str],
        rows: Iterable[Dict[str, object]],
        *,
        filename_pattern: str = "{index:05d}_{field}",
        index_field: str = "id",
        progress_callback: callable | None = None,
        flatten: bool = False,
        template_metadata: Optional[PdfTemplate] = None,
    ) -> List[Path]:
        """Fill a PDF for each row and write the results to ``destination_dir``."""
        template_path = template_path.expanduser().resolve()
        destination_dir = destination_dir.expanduser().resolve()
        destination_dir.mkdir(parents=True, exist_ok=True)

        if not isinstance(rows, list):
            rows = list(rows)

        base_doc = None
        close_base_doc = False
        if flatten:
            if template_metadata:
                base_doc = template_metadata.document
            else:
                base_doc = fitz.open(template_path)
                close_base_doc = True

        outputs: List[Path] = []
        total_rows = len(rows)
        for index, row in enumerate(rows, start=1):
            payload: Dict[str, str] = {}
            for field_name, column_name in field_mapping.items():
                value = row.get(column_name, "")
                payload[field_name] = "" if value is None else str(value)

            label_value = row.get(index_field) or index
            filename = filename_pattern.format(index=index, field=label_value)
            output_path = destination_dir / f"{filename}.pdf"

            if flatten:
                self._write_flattened_pdf(
                    template_doc=base_doc,
                    template_path=template_path,
                    payload=payload,
                    output_path=output_path,
                )
            else:
                self._write_interactive_pdf(template_path, payload, output_path)

            outputs.append(output_path)
            if progress_callback:
                progress_callback(index, total_rows)

        if close_base_doc and base_doc is not None:
            base_doc.close()

        return outputs

    def _write_interactive_pdf(self, template_path: Path, payload: Dict[str, str], output_path: Path) -> None:
        """Fill the PDF form fields while keeping them editable."""
        template_bytes = template_path.read_bytes()
        reader = PdfReader(io.BytesIO(template_bytes))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        for page in writer.pages:
            writer.update_page_form_field_values(page, payload)

        with output_path.open("wb") as handle:
            writer.write(handle)

    def _write_flattened_pdf(
        self,
        template_doc: fitz.Document | None,
        template_path: Path,
        payload: Dict[str, str],
        output_path: Path,
    ) -> None:
        """Render field values directly onto the PDF and remove form widgets."""
        close_template_doc = False
        src_doc = template_doc
        if src_doc is None:
            src_doc = fitz.open(template_path)
            close_template_doc = True

        working = fitz.open()
        working.insert_pdf(src_doc)

        for page_index in range(working.page_count):
            page = working.load_page(page_index)
            widgets = list(page.widgets() or [])
            for widget in widgets:
                field_name = widget.field_name
                text = payload.get(field_name, "")
                if text:
                    point = fitz.Point(widget.rect.x0 + 2, widget.rect.y1 - 4)
                    page.insert_text(
                        point,
                        text,
                        fontname="helv",
                        fontsize=11,
                    )
                page.delete_widget(widget)

        try:
            working.delete_pdf_form()
        except AttributeError:
            pass

        working.save(output_path)
        working.close()

        if close_template_doc and src_doc is not None:
            src_doc.close()

        self._strip_acroform(output_path)

    def _strip_acroform(self, pdf_path: Path) -> None:
        """Remove residual AcroForm metadata from a PDF."""
        reader = PdfReader(str(pdf_path))
        writer = PdfWriter()
        for page in reader.pages:
            page.pop(NameObject("/Annots"), None)
            writer.add_page(page)

        root = writer._root_object  # type: ignore[attr-defined]
        root.pop(NameObject("/AcroForm"), None)

        with pdf_path.open("wb") as handle:
            writer.write(handle)
