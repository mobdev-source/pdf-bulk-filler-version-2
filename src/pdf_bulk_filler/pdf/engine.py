"""PDF utilities for loading templates and generating filled documents."""

from __future__ import annotations

from dataclasses import dataclass
import io
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional

import fitz  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import BooleanObject, DictionaryObject, NameObject, NumberObject

from pdf_bulk_filler.mapping.rules import coerce_rules, evaluate_rules


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
        rule_spec: Any,
        rows: Iterable[Dict[str, object]],
        *,
        filename_pattern: str = "{index:05d}_{field}",
        index_field: str = "id",
        filename_builder: Callable[[Mapping[str, Any], int], str] | None = None,
        progress_callback: callable | None = None,
        flatten: bool = False,
        template_metadata: Optional[PdfTemplate] = None,
        read_only: bool = False,
    ) -> List[Path]:
        """Fill a PDF for each row and write the results to ``destination_dir``."""
        template_path = template_path.expanduser().resolve()
        destination_dir = destination_dir.expanduser().resolve()
        destination_dir.mkdir(parents=True, exist_ok=True)

        rules = coerce_rules(rule_spec)

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
            row_mapping: Mapping[str, Any]
            if isinstance(row, Mapping):
                row_mapping = row
            else:
                row_mapping = dict(row)

            payload = evaluate_rules(rules, row_mapping)

            label_value = row_mapping.get(index_field) or index
            filename_value: str = ""
            if filename_builder is not None:
                try:
                    filename_value = str(filename_builder(row_mapping, index)).strip()
                except Exception:
                    filename_value = ""
            if not filename_value:
                filename_value = str(filename_pattern.format(index=index, field=label_value))
            sanitized = filename_value.replace("/", "_").replace("\\", "_").strip()
            if not sanitized:
                sanitized = f"{index:05d}"
            output_path = destination_dir / f"{sanitized}.pdf"

            if flatten:
                self._write_flattened_pdf(
                    template_doc=base_doc,
                    template_path=template_path,
                    payload=payload,
                    output_path=output_path,
                )
            else:
                self._write_interactive_pdf(
                    template_path,
                    payload,
                    output_path,
                    read_only=read_only,
                )

            outputs.append(output_path)
            if progress_callback:
                progress_callback(index, total_rows)

        if close_base_doc and base_doc is not None:
            base_doc.close()

        return outputs

    def _write_interactive_pdf(
        self,
        template_path: Path,
        payload: Dict[str, object],
        output_path: Path,
        *,
        read_only: bool = False,
    ) -> None:
        """Fill the PDF form fields while keeping them editable."""
        template_bytes = template_path.read_bytes()
        reader = PdfReader(io.BytesIO(template_bytes))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        text_updates: Dict[str, object] = {}
        checkbox_updates: Dict[str, Any] = {}
        for field_name, value in payload.items():
            kind, normalized = self._normalize_payload_value(value)
            if kind == "checkbox":
                checkbox_updates[field_name] = normalized
            else:
                text_updates[field_name] = normalized

        if text_updates:
            for page in writer.pages:
                writer.update_page_form_field_values(page, text_updates)

        for page in writer.pages:
            annotations_obj = page.get("/Annots")
            if not annotations_obj:
                continue
            if hasattr(annotations_obj, "get_object"):
                annotations_obj = annotations_obj.get_object()
            if not annotations_obj:
                continue
            annotation_refs = list(annotations_obj)

            if checkbox_updates:
                for annotation_ref in annotation_refs:
                    annotation = annotation_ref.get_object()
                    field_name = annotation.get("/T")
                    if not field_name or field_name not in checkbox_updates:
                        continue
                    resolved_state = self._resolve_checkbox_state(
                        checkbox_updates[field_name], annotation
                    )
                    annotation.update(
                        {
                            NameObject("/V"): resolved_state,
                            NameObject("/AS"): resolved_state,
                        }
                    )

            if read_only:
                for annotation_ref in annotation_refs:
                    annotation = annotation_ref.get_object()
                    flags = int(annotation.get("/Ff", 0))
                    annotation.update({NameObject("/Ff"): NumberObject(flags | 1)})

        acro_form_obj = writer._root_object.get(NameObject("/AcroForm"))
        if acro_form_obj is None:
            acro_form = DictionaryObject()
            writer._root_object.update({NameObject("/AcroForm"): acro_form})
        else:
            acro_form = acro_form_obj.get_object() if hasattr(acro_form_obj, "get_object") else acro_form_obj
        acro_form.update({NameObject("/NeedAppearances"): BooleanObject(True)})

        with output_path.open("wb") as handle:
            writer.write(handle)

    @staticmethod
    def _normalize_payload_value(value: Any) -> tuple[str, Any]:
        """Return the type (checkbox/text) and normalized value."""
        if isinstance(value, bool):
            return "checkbox", value
        if value is None:
            return "text", ""
        if isinstance(value, (int, float)):
            return "text", str(value)
        if isinstance(value, str):
            stripped = value.strip()
            if stripped.startswith("/") and len(stripped) > 1:
                return "checkbox", stripped
            lowered = stripped.lower()
            if lowered in {"yes", "no", "on", "off", "true", "false", "1", "0", "checked", "unchecked"}:
                return "checkbox", stripped
            return "text", stripped
        return "text", str(value)

    @staticmethod
    def _resolve_checkbox_state(value: Any, annotation: DictionaryObject) -> NameObject:
        """Resolve the correct checkbox state name."""
        def _available_states() -> list[NameObject]:
            ap = annotation.get("/AP")
            if hasattr(ap, "get_object"):
                ap = ap.get_object()
            if isinstance(ap, DictionaryObject):
                normal = ap.get("/N")
                if hasattr(normal, "get_object"):
                    normal = normal.get_object()
                if isinstance(normal, DictionaryObject):
                    return [state for state in normal.keys() if isinstance(state, NameObject)]
            return []

        def _match_state(target: NameObject) -> NameObject | None:
            for state in states:
                if state == target or state[1:].lower() == target[1:].lower():
                    return state
            return None

        def _first_on_state() -> NameObject | None:
            for state in states:
                if state != NameObject("/Off"):
                    return state
            return None

        def _first_off_state() -> NameObject | None:
            for state in states:
                lowered = state[1:].lower() if len(state) > 1 else ""
                if state == NameObject("/Off") or lowered in {"off", "no", "false", "unchecked", "0"}:
                    return state
            return None

        states = _available_states()

        if isinstance(value, NameObject):
            match = _match_state(value)
            return match or value
        if isinstance(value, str):
            stripped = value.strip()
            if stripped.startswith("/") and len(stripped) > 1:
                target = NameObject(stripped)
                match = _match_state(target)
                return match or target
            lowered = stripped.lower()
            if lowered in {"yes", "true", "on", "1", "checked"}:
                target = NameObject("/Yes")
                match = _match_state(target)
                if match:
                    return match
                fallback_on = _first_on_state()
                return fallback_on or target
            if lowered in {"no", "false", "off", "0", "unchecked"}:
                candidates = [NameObject("/Off")]
                if stripped:
                    candidates.append(NameObject(f"/{stripped}"))
                for candidate in candidates:
                    match = _match_state(candidate)
                    if match:
                        return match
                fallback_off = _first_off_state()
                return fallback_off or NameObject("/Off")
            if stripped:
                target = NameObject(f"/{stripped}")
                match = _match_state(target)
                return match or target
            return NameObject("/Off")
        if isinstance(value, bool):
            target = NameObject("/Yes" if value else "/Off")
            match = _match_state(target)
            if match:
                return match
            return _first_on_state() if value else (_first_off_state() or target)
        if value:
            target = NameObject("/Yes")
            match = _match_state(target)
            if match:
                return match
            fallback = _first_on_state()
            return fallback or target
        return _match_state(NameObject("/Off")) or NameObject("/Off")

    def _write_flattened_pdf(
        self,
        template_doc: fitz.Document | None,
        template_path: Path,
        payload: Dict[str, object],
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
                if isinstance(text, str) and text.startswith("/"):
                    lowered = text.lower()
                    if lowered in {"/off", "/no"}:
                        text = ""
                    else:
                        text = "X"
                elif isinstance(text, bool):
                    text = "X" if text else ""
                elif isinstance(text, str):
                    lowered = text.strip().lower()
                    if lowered in {"yes", "true", "on", "1", "checked"}:
                        text = "X"
                    elif lowered in {"no", "false", "off", "0", "unchecked"}:
                        text = ""
                    else:
                        text = text.strip()
                else:
                    text = "" if text is None else str(text)
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
