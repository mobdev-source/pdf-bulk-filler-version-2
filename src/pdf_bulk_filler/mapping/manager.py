"""Persistence helpers for column-to-field mapping configurations."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Mapping, MutableMapping


@dataclass
class MappingModel:
    """In-memory representation of column-field relationships."""

    source_data: Path | None = None
    pdf_template: Path | None = None
    data_sheet: str | None = None
    header_row: int | None = None
    data_row: int | None = None
    column_offset: int | None = None
    assignments: MutableMapping[str, str] = field(default_factory=dict)

    def assign(self, field_name: str, column_name: str) -> None:
        """Associate a PDF field with a data column."""
        self.assignments[field_name] = column_name

    def remove(self, field_name: str) -> None:
        """Remove an association for the given PDF field."""
        self.assignments.pop(field_name, None)

    def resolve(self, field_name: str) -> str | None:
        """Return the column mapped to the given field, if present."""
        return self.assignments.get(field_name)


class MappingManager:
    """Serialize and hydrate mapping models from JSON files."""

    def save(self, destination: Path, mapping: MappingModel) -> None:
        payload = {
            "source_data": str(mapping.source_data) if mapping.source_data else None,
            "pdf_template": str(mapping.pdf_template) if mapping.pdf_template else None,
            "data_sheet": mapping.data_sheet,
            "header_row": mapping.header_row,
            "data_row": mapping.data_row,
            "column_offset": mapping.column_offset,
            "assignments": dict(mapping.assignments),
        }
        destination = destination.expanduser().resolve()
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def load(self, source: Path) -> MappingModel:
        raw = json.loads(Path(source).read_text(encoding="utf-8"))
        return MappingModel(
            source_data=Path(raw["source_data"]) if raw.get("source_data") else None,
            pdf_template=Path(raw["pdf_template"]) if raw.get("pdf_template") else None,
            data_sheet=raw.get("data_sheet"),
            header_row=raw.get("header_row"),
            data_row=raw.get("data_row"),
            column_offset=raw.get("column_offset"),
            assignments=dict(raw.get("assignments", {})),
        )

    @staticmethod
    def mapping_to_fields(mapping: MappingModel) -> Dict[str, str]:
        """Return a standard dictionary for PDF filling utilities."""
        return dict(mapping.assignments)
