"""Rule definitions and evaluation helpers for PDF field mappings."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, Iterable, Mapping, MutableMapping

import datetime as dt
import pandas as pd


class RuleType(str, Enum):
    """Supported mapping strategies."""

    VALUE = "value"
    LITERAL = "literal"
    CHOICE = "choice"
    CONCAT = "concat"


def _ensure_targets(rule_name: str, targets: Iterable[str] | None) -> list[str]:
    resolved = [target for target in targets or [] if target]
    if resolved:
        return resolved
    return [rule_name]


@dataclass
class MappingRule:
    """Declarative rule describing how to populate one or more PDF fields."""

    name: str
    rule_type: RuleType | str = RuleType.VALUE
    targets: list[str] = field(default_factory=list)
    options: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if isinstance(self.rule_type, str):
            try:
                self.rule_type = RuleType(self.rule_type)
            except ValueError as exc:
                raise ValueError(f"Unknown rule type '{self.rule_type}'") from exc
        self.options = dict(self.options or {})
        self.targets = _ensure_targets(self.name, self.targets)

    @classmethod
    def from_direct_column(cls, field_name: str, column_name: str) -> "MappingRule":
        """Return a rule mirroring the legacy direct column mapping."""
        return cls(
            name=field_name,
            rule_type=RuleType.VALUE,
            targets=[field_name],
            options={"column": column_name},
        )

    @classmethod
    def from_json(cls, payload: Mapping[str, Any]) -> "MappingRule":
        """Create a rule from a serialized JSON payload."""
        rule_type = RuleType(payload.get("type", RuleType.VALUE))
        name = str(payload.get("name") or payload.get("field") or "")
        targets = list(payload.get("targets") or [])
        options = dict(payload.get("options") or {})
        if not name:
            raise ValueError("Mapping rule is missing a name.")
        return cls(name=name, rule_type=rule_type, targets=targets, options=options)

    def to_json(self) -> Dict[str, Any]:
        """Return a JSON-friendly representation."""
        return {
            "name": self.name,
            "type": self.type_enum().value,
            "targets": list(self.targets),
            "options": dict(self.options),
        }

    def describe(self) -> str:
        """Return a compact human-friendly summary for UI tables."""
        rule_type = self.type_enum()
        if rule_type is RuleType.VALUE:
            column = self.options.get("column", "<missing>")
            return f"{column}"
        if rule_type is RuleType.LITERAL:
            value = self.options.get("value", "")
            return f"Literal: {value}"
        if rule_type is RuleType.CHOICE:
            source = self.options.get("source", "<source>")
            cases = self.options.get("cases", {})
            case_count = len(cases)
            return f"Choice from {source} ({case_count} cases)"
        if rule_type is RuleType.CONCAT:
            columns = self.options.get("columns", [])
            return " + ".join(str(col) for col in columns) or "Concatenate"
        return rule_type.value

    def type_enum(self) -> RuleType:
        """Return the rule type as a ``RuleType`` enum."""
        if isinstance(self.rule_type, RuleType):
            return self.rule_type
        try:
            self.rule_type = RuleType(self.rule_type)
        except ValueError as exc:
            raise ValueError(f"Unknown rule type '{self.rule_type}'") from exc
        return self.rule_type


class RuleEvaluator:
    """Compute PDF payload fragments for a rule and a single data row."""

    def evaluate(self, rule: MappingRule, row: Mapping[str, Any]) -> Dict[str, str]:
        """Return field values produced by ``rule``."""
        rule_type = rule.rule_type
        if isinstance(rule_type, str):
            try:
                rule_type = RuleType(rule_type)
            except ValueError as exc:
                raise ValueError(f"Unsupported rule type: {rule_type}") from exc
        if rule_type is RuleType.VALUE:
            return self._evaluate_value(rule, row)
        if rule_type is RuleType.LITERAL:
            return self._evaluate_literal(rule)
        if rule_type is RuleType.CHOICE:
            return self._evaluate_choice(rule, row)
        if rule_type is RuleType.CONCAT:
            return self._evaluate_concat(rule, row)
        raise ValueError(f"Unsupported rule type: {rule_type}")

    @staticmethod
    def _stringify(value: Any, *, default: str = "") -> str:
        if value is None:
            return default
        if isinstance(value, str):
            return value
        try:
            if pd.isna(value):  # type: ignore[arg-type]
                return default
        except TypeError:
            pass
        if isinstance(value, pd.Timestamp):
            if pd.isna(value):
                return default
            try:
                value = value.to_pydatetime()
            except AttributeError:
                pass
            except ValueError:
                pass
        if isinstance(value, dt.datetime):
            if value.time() == dt.time(0, 0, 0):
                return value.strftime("%Y-%m-%d")
            return value.strftime("%Y-%m-%d %H:%M:%S")
        if isinstance(value, dt.date):
            return value.strftime("%Y-%m-%d")
        try:
            if pd.isna(value):  # type: ignore[arg-type]
                return default
        except TypeError:
            pass
        return str(value)

    @staticmethod
    def _prepare_for_format(value: Any) -> Any:
        if isinstance(value, pd.Timestamp):
            if pd.isna(value):
                return None
            try:
                return value.to_pydatetime()
            except AttributeError:
                pass
            except ValueError:
                return None
        if isinstance(value, dt.datetime):
            return value
        if isinstance(value, dt.date):
            return value
        return value

    def _evaluate_value(self, rule: MappingRule, row: Mapping[str, Any]) -> Dict[str, str]:
        options = rule.options
        column = options.get("column")
        default = options.get("default", "")
        targets = rule.targets
        if not column:
            return {target: default for target in targets}
        raw = row.get(column, default)
        template = options.get("format")
        if isinstance(template, str) and template:
            try:
                context = {key: self._prepare_for_format(value) for key, value in row.items()}
                formatted = template.format(
                    value=self._prepare_for_format(raw),
                    **context,
                )
            except Exception:  # noqa: BLE001
                formatted = self._stringify(raw, default=default)
            else:
                return {target: str(formatted) for target in targets}

        text = self._stringify(raw, default=default)
        return {target: text for target in targets}

    def _evaluate_literal(self, rule: MappingRule) -> Dict[str, str]:
        value = self._stringify(rule.options.get("value", ""), default="")
        return {target: value for target in rule.targets}

    def _evaluate_choice(self, rule: MappingRule, row: Mapping[str, Any]) -> Dict[str, str]:
        options = rule.options
        source = options.get("source")
        cases = options.get("cases", {})
        default = options.get("default", {})
        raw_value = row.get(source) if source else None

        if raw_value in cases:
            selected = cases[raw_value]
        elif str(raw_value) in cases:
            selected = cases[str(raw_value)]
        else:
            selected = default

        return self._normalize_choice_output(rule, selected)

    def _normalize_choice_output(self, rule: MappingRule, selected: Any) -> Dict[str, str]:
        if isinstance(selected, Mapping):
            result: Dict[str, str] = {}
            for key, value in selected.items():
                result[str(key)] = self._stringify(value, default="")
            return result
        text = self._stringify(selected, default="")
        return {target: text for target in rule.targets}

    def _evaluate_concat(self, rule: MappingRule, row: Mapping[str, Any]) -> Dict[str, str]:
        options = rule.options
        columns: Iterable[str] = options.get("columns", [])
        separator: str = options.get("separator", ", ")
        skip_empty: bool = bool(options.get("skip_empty", True))
        default = options.get("default", "")

        parts = []
        for column in columns:
            value = row.get(column)
            text = self._stringify(value, default="")
            if skip_empty and not text:
                continue
            parts.append(text)

        combined = separator.join(parts) if parts else default
        prefix = options.get("prefix", "")
        suffix = options.get("suffix", "")
        if prefix:
            combined = f"{prefix}{combined}"
        if suffix:
            combined = f"{combined}{suffix}"

        return {target: combined for target in rule.targets}


def evaluate_rules(
    rules: Iterable[MappingRule],
    row: Mapping[str, Any],
) -> Dict[str, str]:
    """Evaluate every rule and merge their outputs."""
    evaluator = RuleEvaluator()
    payload: Dict[str, str] = {}
    for rule in rules:
        fragment = evaluator.evaluate(rule, row)
        payload.update(fragment)
    return payload


def coerce_rules(spec: Any) -> list[MappingRule]:
    """Return a list of mapping rules from various legacy inputs."""
    if spec is None:
        return []
    if isinstance(spec, MappingRule):
        return [spec]
    if hasattr(spec, "iter_rules"):
        try:
            rules_iter = spec.iter_rules()  # type: ignore[attr-defined]
        except TypeError:
            pass
        else:
            return list(rules_iter)
    if isinstance(spec, Mapping):
        result: list[MappingRule] = []
        for field_name, value in spec.items():
            if isinstance(value, MappingRule):
                rule = value
            else:
                rule = MappingRule.from_direct_column(str(field_name), str(value))
            rule.name = str(field_name)
            if not rule.targets:
                rule.targets = [rule.name]
            result.append(rule)
        return result
    if isinstance(spec, Iterable) and not isinstance(spec, (str, bytes)):
        result = []
        for item in spec:
            if isinstance(item, MappingRule):
                result.append(item)
                continue
            if isinstance(item, tuple) and len(item) == 2:
                field_name, column = item
                result.append(
                    MappingRule.from_direct_column(str(field_name), str(column))
                )
                continue
            raise TypeError(f"Unsupported rule specification: {item!r}")
        return result
    raise TypeError(f"Unsupported rule specification: {spec!r}")


def rules_from_legacy(assignments: Mapping[str, str]) -> MutableMapping[str, MappingRule]:
    """Convert a legacy field->column map into modern rules."""
    converted: MutableMapping[str, MappingRule] = {}
    for field_name, column_name in assignments.items():
        converted[field_name] = MappingRule.from_direct_column(field_name, column_name)
    return converted
