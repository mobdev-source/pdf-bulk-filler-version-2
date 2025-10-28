import json
from pathlib import Path
from datetime import datetime

import pandas as pd

from pdf_bulk_filler.mapping.manager import MappingManager, MappingModel
from pdf_bulk_filler.mapping.rules import MappingRule, RuleType, RuleEvaluator, coerce_rules, evaluate_rules


def test_rule_evaluator_value_handles_missing_values():
    row = {"First Name": "Alice", "Last Name": None}
    rule = MappingRule.from_direct_column("FullName", "First Name")
    evaluator = RuleEvaluator()

    result = evaluator.evaluate(rule, row)

    assert result == {"FullName": "Alice"}


def test_rule_evaluator_value_formats_datetime_without_time():
    rule = MappingRule.from_direct_column("Birthdate", "BirthDate")
    evaluator = RuleEvaluator()
    row = {"BirthDate": pd.Timestamp("1995-12-12 00:00:00")}

    result = evaluator.evaluate(rule, row)

    assert result == {"Birthdate": "1995-12-12"}


def test_rule_evaluator_value_applies_format_string():
    rule = MappingRule(
        name="Birthdate",
        rule_type=RuleType.VALUE,
        targets=["Birthdate"],
        options={"column": "BirthDate", "format": "{value:%Y-%m-%d}"},
    )
    evaluator = RuleEvaluator()
    row = {"BirthDate": datetime(1995, 12, 12, 15, 30)}

    result = evaluator.evaluate(rule, row)

    assert result == {"Birthdate": "1995-12-12"}


def test_rule_evaluator_choice_emits_multiple_fields():
    row = {"Gender": "F"}
    rule = MappingRule(
        name="Gender",
        rule_type=RuleType.CHOICE,
        targets=["Male", "Female"],
        options={
            "source": "Gender",
            "cases": {
                "M": {"Male": "Yes", "Female": ""},
                "F": {"Male": "", "Female": "Yes"},
            },
            "default": {"Male": "", "Female": ""},
        },
    )
    evaluator = RuleEvaluator()

    result = evaluator.evaluate(rule, row)

    assert result == {"Male": "", "Female": "Yes"}


def test_rule_evaluator_choice_supports_column_actions():
    row = {"Status": "Other", "Other Description": "Separated"}
    rule = MappingRule(
        name="Status",
        rule_type=RuleType.CHOICE,
        targets=["Single", "Married", "Other", "OtherText"],
        options={
            "source": "Status",
            "cases": {
                "Other": {
                    "Other": {"mode": "checkbox", "checked": True},
                    "OtherText": {"mode": "column", "column": "Other Description"},
                },
                "Single": {
                    "Single": {"mode": "checkbox", "checked": True},
                    "Married": {"mode": "checkbox", "checked": False},
                },
            },
        },
    )
    evaluator = RuleEvaluator()

    result = evaluator.evaluate(rule, row)

    assert result["Other"] is True
    assert result["OtherText"] == "Separated"


def test_rule_evaluator_concat_skips_empty_segments():
    row = {"Street": "Blk 58 Lot 8 Annapolis St", "Barangay": "", "City": "Buendia"}
    rule = MappingRule(
        name="Address",
        rule_type=RuleType.CONCAT,
        targets=["Address"],
        options={
            "columns": ["Street", "Barangay", "City"],
            "separator": ", ",
            "skip_empty": True,
        },
    )
    evaluator = RuleEvaluator()

    result = evaluator.evaluate(rule, row)

    assert result == {"Address": "Blk 58 Lot 8 Annapolis St, Buendia"}


def test_evaluate_rules_merges_outputs():
    row = {"First": "Ana", "Last": "Dela Cruz", "Status": "Married"}
    rules = [
        MappingRule.from_direct_column("FirstName", "First"),
        MappingRule(
            name="Status",
            rule_type=RuleType.CHOICE,
            targets=["Single", "Married", "Other", "OtherText"],
            options={
                "source": "Status",
                "cases": {
                    "Single": {"Single": "Yes", "Married": "", "Other": ""},
                    "Married": {"Single": "", "Married": "Yes", "Other": ""},
                    "Widow": {
                        "Single": "",
                        "Married": "",
                        "Other": "Yes",
                        "OtherText": "Widow",
                    },
                },
                "default": {"Single": "", "Married": "", "Other": "", "OtherText": ""},
            },
        ),
    ]

    payload = evaluate_rules(rules, row)

    assert payload["FirstName"] == "Ana"
    assert payload["Single"] == ""
    assert payload["Married"] == "Yes"
    assert payload["Other"] == ""
    assert payload.get("OtherText", "") == ""


def test_mapping_manager_loads_legacy_assignments(tmp_path: Path):
    payload = {
        "source_data": "data.xlsx",
        "pdf_template": "template.pdf",
        "assignments": {"NameField": "Full Name", "AgeField": "Age"},
    }
    path = tmp_path / "legacy.json"
    path.write_text(json.dumps(payload), encoding="utf-8")

    manager = MappingManager()
    model = manager.load(path)

    assert set(model.rules.keys()) == {"NameField", "AgeField"}
    assert model.resolve("NameField").options["column"] == "Full Name"


def test_mapping_manager_roundtrip_with_rules(tmp_path: Path):
    model = MappingModel()
    model.assign(
        "GenderRule",
        MappingRule(
            name="GenderRule",
            rule_type=RuleType.CHOICE,
            targets=["Male", "Female"],
            options={
                "source": "Gender",
                "cases": {"M": {"Male": "Yes", "Female": ""}},
                "default": {"Male": "", "Female": "Yes"},
            },
        ),
    )
    destination = tmp_path / "rules.json"

    manager = MappingManager()
    manager.save(destination, model)

    loaded = manager.load(destination)
    assert "GenderRule" in loaded.rules
    restored = loaded.resolve("GenderRule")
    assert restored.rule_type is RuleType.CHOICE
    assert restored.options["source"] == "Gender"


def test_mapping_model_assign_shares_choice_rule_across_targets():
    model = MappingModel()
    rule = MappingRule(
        name="Female",
        rule_type=RuleType.CHOICE,
        targets=["Female", "Male"],
        options={
            "source": "Gender",
            "cases": [{"match": "F", "outputs": {"Female": "Yes", "Male": ""}}],
            "default": {"Female": "", "Male": ""},
        },
    )

    model.assign("Female", rule)

    assert set(model.rules.keys()) == {"Female", "Male"}
    female = model.resolve("Female")
    male = model.resolve("Male")
    assert female is not male
    assert female.targets == ["Female", "Male"]
    assert male.targets == ["Female", "Male"]
    assert male.options is not female.options
    assert male.options == female.options
    assert male.options.get("case_map", {})["F"]["Female"] == "Yes"


def test_mapping_model_assign_prunes_removed_targets():
    model = MappingModel()
    model.assign(
        "Female",
        MappingRule(
            name="Female",
            rule_type=RuleType.CHOICE,
            targets=["Female", "Male"],
            options={"source": "Gender", "cases": {"F": {"Female": "Yes"}}},
        ),
    )

    model.assign(
        "Female",
        MappingRule(
            name="Female",
            rule_type=RuleType.CHOICE,
            targets=["Female"],
            options={"source": "Gender", "cases": {"F": {"Female": "Yes"}}},
        ),
    )

    assert "Female" in model.rules
    assert "Male" not in model.rules

def test_mapping_rule_accepts_string_types():
    rule = MappingRule(name="Address", rule_type="concat", targets=["Address"], options={"columns": ["Street"]})
    evaluator = RuleEvaluator()

    result = evaluator.evaluate(rule, {"Street": "123 Main"})

    assert result == {"Address": "123 Main"}
    assert rule.type_enum() is RuleType.CONCAT


def test_coerce_rules_handles_iterable_of_tuples():
    rules = coerce_rules([("FullName", "Name"), ("AgeField", "Age")])
    names = {rule.name for rule in rules}

    assert names == {"FullName", "AgeField"}
    assert all(rule.type_enum() is RuleType.VALUE for rule in rules)
