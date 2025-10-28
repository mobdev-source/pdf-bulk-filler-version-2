"""Persistence helpers for column-to-field mapping configurations."""

from __future__ import annotations

import copy
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, Mapping, MutableMapping

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
        previous = self.rules.get(field_name)

        if isinstance(rule, MappingRule):
            normalized = MappingRule(
                name=field_name,
                rule_type=rule.rule_type,
                targets=list(rule.targets),
                options=copy.deepcopy(rule.options),
            )
        else:
            normalized = MappingRule.from_direct_column(field_name, rule)

        normalized.name = field_name
        targets = list(dict.fromkeys(t for t in (normalized.targets or []) if t))
        if field_name not in targets:
            targets.insert(0, field_name)
        normalized.targets = targets or [field_name]

        if normalized.type_enum() is RuleType.CHOICE:
            self._ensure_choice_case_map(normalized.options)

        previous_targets: set[str] = set(previous.targets) if previous else set()
        new_target_set = set(normalized.targets)
        stale_targets = previous_targets - new_target_set
        for stale in stale_targets:
            if stale == field_name:
                continue
            existing = self.rules.get(stale)
            if existing and set(existing.targets) == previous_targets:
                self.rules.pop(stale, None)

        # Assign clones for every targeted field so the dialog shows shared rules regardless of entry point.
        for target in normalized.targets:
            clone = MappingRule(
                name=target,
                rule_type=normalized.rule_type,
                targets=list(normalized.targets),
                options=copy.deepcopy(normalized.options),
            )
            self.rules[target] = clone

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

    @staticmethod
    def _ensure_choice_case_map(options: Dict[str, object]) -> None:
        existing = options.get("case_map")
        if isinstance(existing, Mapping) and existing:
            return

        cases = options.get("cases")
        case_map: Dict[str, Dict[str, object]] = {}

        if isinstance(cases, Mapping):
            for key, outputs in cases.items():
                if isinstance(outputs, Mapping):
                    case_map[str(key)] = copy.deepcopy(outputs)
        elif isinstance(cases, Iterable) and not isinstance(cases, (str, bytes)):
            for case in cases:
                if not isinstance(case, Mapping):
                    continue
                match_value = str(case.get("match", "")).strip()
                if not match_value:
                    continue
                outputs = case.get("outputs")
                if isinstance(outputs, Mapping):
                    case_map[match_value] = copy.deepcopy(outputs)

        if case_map:
            options["case_map"] = case_map


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
