"""Persistence helpers for column-to-field mapping configurations."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, MutableMapping

from pdf_bulk_filler.mapping.rules import MappingRule, RuleType, rules_from_legacy


@dataclass
class MappingModel:
    """In-memory representation of mapping rules."""

    source_data: Path | None = None
    pdf_template: Path | None = None
    data_sheet: str | None = None
    header_row: int | None = None
    data_row: int | None = None
    column_offset: int | None = None
    rules: MutableMapping[str, MappingRule] = field(default_factory=dict)

    def assign(self, field_name: str, rule: MappingRule | str) -> None:
        """Associate a PDF field with a rule or column."""
        if isinstance(rule, MappingRule):
            normalized = rule
        else:
            normalized = MappingRule.from_direct_column(field_name, rule)
        normalized.name = field_name
        normalized.targets = [field_name] if not normalized.targets else normalized.targets
        self.rules[field_name] = normalized

    def remove(self, field_name: str) -> None:
        """Remove an association for the given PDF field."""
        self.rules.pop(field_name, None)

    def resolve(self, field_name: str) -> MappingRule | None:
        """Return the rule mapped to the given field, if present."""
        return self.rules.get(field_name)

    def iter_rules(self) -> Iterable[MappingRule]:
        """Yield rules in insertion order."""
        return self.rules.values()

    def to_legacy_fields(self) -> Dict[str, str]:
        """Return a simplified mapping for legacy consumers."""
        legacy: Dict[str, str] = {}
        for field_name, rule in self.rules.items():
            if rule.type_enum() is RuleType.VALUE:
                column = rule.options.get("column")
                if column:
                    legacy[field_name] = column
        return legacy

    @property
    def assignments(self) -> MutableMapping[str, MappingRule]:
        """Backward-compatible accessor exposing the internal rule map."""
        return self.rules


class MappingManager:
    """Serialize and hydrate mapping models from JSON files."""

    def save(self, destination: Path, mapping: MappingModel) -> None:
        payload = {
            "version": 2,
            "source_data": str(mapping.source_data) if mapping.source_data else None,
            "pdf_template": str(mapping.pdf_template) if mapping.pdf_template else None,
            "data_sheet": mapping.data_sheet,
            "header_row": mapping.header_row,
            "data_row": mapping.data_row,
            "column_offset": mapping.column_offset,
            "rules": [rule.to_json() for rule in mapping.rules.values()],
        }
        destination = destination.expanduser().resolve()
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def load(self, source: Path) -> MappingModel:
        raw = json.loads(Path(source).read_text(encoding="utf-8"))
        rules_payload = raw.get("rules")
        if isinstance(rules_payload, list):
            rules: Dict[str, MappingRule] = {}
            for rule_data in rules_payload:
                rule = MappingRule.from_json(rule_data)
                rules[rule.name] = rule
        else:
            assignments = dict(raw.get("assignments", {}))
            rules = rules_from_legacy(assignments)
        return MappingModel(
            source_data=Path(raw["source_data"]) if raw.get("source_data") else None,
            pdf_template=Path(raw["pdf_template"]) if raw.get("pdf_template") else None,
            data_sheet=raw.get("data_sheet"),
            header_row=raw.get("header_row"),
            data_row=raw.get("data_row"),
            column_offset=raw.get("column_offset"),
            rules=rules,
        )

    @staticmethod
    def mapping_to_fields(mapping: MappingModel) -> Dict[str, str]:
        """Return a legacy dictionary for PDF filling helpers."""
        return mapping.to_legacy_fields()
