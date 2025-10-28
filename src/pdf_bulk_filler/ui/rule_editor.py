"""Dialog for configuring mapping rules via the GUI."""

from __future__ import annotations

from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional, Sequence

from PySide6 import QtCore, QtGui, QtWidgets

from pdf_bulk_filler.mapping.rules import MappingRule, RuleType

RULE_HELP_TEXT: dict[RuleType, str] = {
    RuleType.VALUE: (
        "Fill the selected PDF fields with the value from a single data column. "
        "Optionally provide a default when the column is empty and a format pattern "
        "(e.g. '{value} USD')."
    ),
    RuleType.LITERAL: (
        "Always write the literal text you provide below. Use this for fixed captions or "
        "checkbox values that never change."
    ),
    RuleType.CHOICE: (
        "Pick a source column, then define how each possible value should populate one or "
        "more PDF fields. Useful for toggling checkboxes (Male/Female) or setting 'Other' text."
    ),
    RuleType.CONCAT: (
        "Combine multiple columns into a single output. Tick the columns to include, drag them "
        "to change order, and choose a separator. Empty values are skipped by default."
    ),
}


class _AutoSizingStack(QtWidgets.QStackedWidget):
    """Stacked widget that resizes to the currently visible page."""

    def sizeHint(self) -> QtCore.QSize:  # type: ignore[override]
        current = self.currentWidget()
        if current is not None:
            return current.sizeHint()
        return super().sizeHint()

    def minimumSizeHint(self) -> QtCore.QSize:  # type: ignore[override]
        current = self.currentWidget()
        if current is not None:
            return current.minimumSizeHint()
        return super().minimumSizeHint()


class _TargetsSelector(QtWidgets.QListWidget):
    """Checkbox list for selecting PDF targets."""

    selectionChanged = QtCore.Signal()

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.itemChanged.connect(lambda _item: self.selectionChanged.emit())

    def set_targets(self, targets: Sequence[str], selected: Iterable[str]) -> None:
        existing = {item.text(): item for item in (self.item(i) for i in range(self.count()))}
        selected_set = set(selected)
        self.clear()
        for target in targets:
            item = existing.get(target, QtWidgets.QListWidgetItem(target))
            item.setText(target)
            item.setFlags(
                QtCore.Qt.ItemIsEnabled
                | QtCore.Qt.ItemIsUserCheckable
                | QtCore.Qt.ItemIsSelectable
            )
            item.setCheckState(QtCore.Qt.Checked if target in selected_set else QtCore.Qt.Unchecked)
            self.addItem(item)

    def selected_targets(self) -> list[str]:
        result: list[str] = []
        for index in range(self.count()):
            item = self.item(index)
            if item.checkState() == QtCore.Qt.Checked:
                result.append(item.text())
        return result


class _ValueConfigWidget(QtWidgets.QWidget):
    """Configuration panel for direct column mapping."""

    def __init__(self, columns: Sequence[str], parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._columns = list(columns)
        self._column_combo = QtWidgets.QComboBox()
        self._default_edit = QtWidgets.QLineEdit()
        self._format_edit = QtWidgets.QLineEdit()

        layout = QtWidgets.QFormLayout(self)
        self._column_combo.addItems(self._columns)
        layout.addRow("Source column:", self._column_combo)
        layout.addRow("Default value:", self._default_edit)
        layout.addRow("Format pattern:", self._format_edit)

    def set_columns(self, columns: Sequence[str]) -> None:
        current = self._column_combo.currentText()
        self._column_combo.blockSignals(True)
        self._column_combo.clear()
        self._column_combo.addItems(columns)
        index = self._column_combo.findText(current)
        if index >= 0:
            self._column_combo.setCurrentIndex(index)
        self._column_combo.blockSignals(False)

    def load_options(self, options: Dict[str, str]) -> None:
        column_value = options.get("column")
        column = column_value if isinstance(column_value, str) else ""
        index = self._column_combo.findText(column)
        if index >= 0:
            self._column_combo.setCurrentIndex(index)
        elif self._column_combo.count():
            self._column_combo.setCurrentIndex(0)
        default_value = options.get("default", "")
        if not isinstance(default_value, str):
            default_value = ""
        format_value = options.get("format", "")
        if not isinstance(format_value, str):
            format_value = ""
        self._default_edit.setText(default_value)
        self._format_edit.setText(format_value)

    def build_options(self) -> Dict[str, str]:
        return {
            "column": self._column_combo.currentText(),
            "default": self._default_edit.text(),
            "format": self._format_edit.text(),
        }

    def current_column(self) -> str:
        return self._column_combo.currentText()


class _LiteralConfigWidget(QtWidgets.QWidget):
    """Configuration panel for literal rule values."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._value_edit = QtWidgets.QLineEdit()
        layout = QtWidgets.QFormLayout(self)
        layout.addRow("Literal value:", self._value_edit)

    def load_options(self, options: Dict[str, str]) -> None:
        self._value_edit.setText(options.get("value", ""))

    def build_options(self) -> Dict[str, str]:
        return {"value": self._value_edit.text()}


class _ConcatConfigWidget(QtWidgets.QWidget):
    """Configuration panel for concatenation rules."""

    def __init__(self, columns: Sequence[str], parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._columns = list(columns)

        self._column_list = QtWidgets.QListWidget()
        self._column_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self._column_list.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self._column_list.viewport().setAcceptDrops(True)
        self._column_list.setDragEnabled(True)

        self._separator_edit = QtWidgets.QLineEdit(", ")
        self._prefix_edit = QtWidgets.QLineEdit()
        self._suffix_edit = QtWidgets.QLineEdit()
        self._skip_empty_check = QtWidgets.QCheckBox("Skip empty values")
        self._skip_empty_check.setChecked(True)

        self._populate_columns()

        form = QtWidgets.QFormLayout()
        form.addRow("Separator:", self._separator_edit)
        form.addRow("Prefix:", self._prefix_edit)
        form.addRow("Suffix:", self._suffix_edit)
        form.addRow("", self._skip_empty_check)

        instructions = QtWidgets.QLabel(
            "Tick the columns to include. Drag entries to change their order in the output."
        )
        instructions.setWordWrap(True)
        instructions.setObjectName("concatInstructions")

        settings_panel = QtWidgets.QWidget()
        settings_layout = QtWidgets.QVBoxLayout(settings_panel)
        settings_layout.setContentsMargins(0, 0, 0, 0)
        settings_layout.setSpacing(6)
        settings_layout.addLayout(form)
        settings_layout.addWidget(instructions)
        settings_layout.addStretch(1)

        layout = QtWidgets.QHBoxLayout(self)
        layout.addWidget(self._column_list, 3)
        layout.addWidget(settings_panel, 4)

    def _populate_columns(self, selected: Iterable[str] | None = None) -> None:
        selected_list = [str(name) for name in selected or [] if name]
        selected_seen: set[str] = set()
        available_set = {str(column) for column in self._columns}

        self._column_list.clear()

        def _add_item(column_name: str, *, checked: bool, enabled: bool = True) -> None:
            item = QtWidgets.QListWidgetItem(column_name)
            flags = QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable
            if enabled:
                flags |= QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsDragEnabled
            item.setFlags(flags)
            item.setCheckState(QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked)
            if not enabled:
                item.setForeground(QtGui.QColor(QtCore.Qt.GlobalColor.gray))
            self._column_list.addItem(item)

        # Preserve the user-defined ordering for selected columns first.
        for column in selected_list:
            if column in selected_seen:
                continue
            if column in available_set:
                _add_item(column, checked=True, enabled=True)
            else:
                _add_item(column, checked=True, enabled=False)
            selected_seen.add(column)

        # Append the remaining available columns in their default order.
        for column in self._columns:
            if column in selected_seen:
                continue
            _add_item(column, checked=column in selected_seen, enabled=True)

    def set_columns(self, columns: Sequence[str]) -> None:
        selected = self.selected_columns()
        self._columns = list(columns)
        self._populate_columns(selected)

    def load_options(self, options: Dict[str, object]) -> None:
        columns = options.get("columns", [])
        if isinstance(columns, list):
            self._populate_columns(columns)
        else:
            self._populate_columns()
        self._separator_edit.setText(str(options.get("separator", ", ")))
        self._prefix_edit.setText(str(options.get("prefix", "")))
        self._suffix_edit.setText(str(options.get("suffix", "")))
        self._skip_empty_check.setChecked(bool(options.get("skip_empty", True)))

    def build_options(self) -> Dict[str, object]:
        return {
            "columns": self.selected_columns(),
            "separator": self._separator_edit.text(),
            "prefix": self._prefix_edit.text(),
            "suffix": self._suffix_edit.text(),
            "skip_empty": self._skip_empty_check.isChecked(),
        }

    def selected_columns(self) -> List[str]:
        columns: List[str] = []
        for index in range(self._column_list.count()):
            item = self._column_list.item(index)
            if item.checkState() == QtCore.Qt.Checked:
                columns.append(item.text())
        return columns


class _ChoiceTargetActionEditor(QtWidgets.QWidget):
    """Editor for configuring how a single PDF target reacts within a case."""

    changed = QtCore.Signal()

    _CHECKBOX_TRUE_VALUES = {"yes", "true", "on", "1", "checked"}
    _CHECKBOX_FALSE_VALUES = {"no", "false", "off", "0", "unchecked"}

    def __init__(self, target: str, columns: Sequence[str], parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._target = target
        self._columns = list(columns)
        self._block_updates = False
        self.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)

        self._mode_combo = QtWidgets.QComboBox()
        self._mode_combo.addItem("Leave unchanged", "ignore")
        self._mode_combo.addItem("Check the box", "checked")
        self._mode_combo.addItem("Uncheck the box", "unchecked")
        self._mode_combo.addItem("Use literal text", "literal")
        self._mode_combo.addItem("Use column value", "column")

        self._text_edit = QtWidgets.QLineEdit()
        self._text_edit.setPlaceholderText("Enter the text to use")

        self._column_combo = QtWidgets.QComboBox()
        self._column_combo.addItems(self._columns)
        self._column_combo.setEditable(False)
        self._column_combo.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
        self._column_combo.setMinimumContentsLength(1)
        self._column_combo.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)

        self._column_fallback = QtWidgets.QLineEdit()
        self._column_fallback.setPlaceholderText("Fallback when column is empty (optional)")
        self._column_fallback.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)

        self._stack = _AutoSizingStack()
        self._stack.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self._stack.addWidget(QtWidgets.QWidget())

        text_page = QtWidgets.QWidget()
        text_layout = QtWidgets.QVBoxLayout(text_page)
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(4)
        text_layout.addWidget(self._text_edit)
        text_page.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self._stack.addWidget(text_page)

        column_page = QtWidgets.QWidget()
        column_layout = QtWidgets.QVBoxLayout(column_page)
        column_layout.setContentsMargins(0, 0, 0, 0)
        column_layout.setSpacing(4)
        column_layout.addWidget(self._column_combo)
        column_layout.addWidget(self._column_fallback)
        column_page.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self._stack.addWidget(column_page)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        layout.addWidget(self._mode_combo)
        layout.addWidget(self._stack)

        self._mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        self._text_edit.textEdited.connect(self._emit_changed)
        self._column_combo.currentIndexChanged.connect(self._emit_changed)
        self._column_fallback.textEdited.connect(self._emit_changed)
        self._on_mode_changed(self._mode_combo.currentIndex())

    def set_columns(self, columns: Sequence[str]) -> None:
        current = self._column_combo.currentText()
        self._columns = list(columns)
        self._column_combo.blockSignals(True)
        self._column_combo.clear()
        self._column_combo.addItems(self._columns)
        index = self._column_combo.findText(current)
        if index >= 0:
            self._column_combo.setCurrentIndex(index)
        self._column_combo.blockSignals(False)

    def load_action(self, action: Any) -> None:
        self._block_updates = True
        mode = "ignore"
        literal_value = ""
        column = ""
        fallback = ""

        if isinstance(action, Mapping):
            mode = str(action.get("mode") or action.get("kind") or action.get("type") or "").lower()
            if not mode:
                if "column" in action:
                    mode = "column"
                elif "checked" in action:
                    mode = "checkbox"
                elif "value" in action:
                    mode = "literal"
            if mode == "column":
                column = str(action.get("column", ""))
                fallback = str(action.get("fallback", ""))
            elif mode in {"literal", "text", "value"}:
                mode = "literal"
                literal_value = str(action.get("value", ""))
            elif mode == "checkbox":
                mode = "checked" if bool(action.get("checked", True)) else "unchecked"
            elif mode == "raw":
                mode = "literal"
                literal_value = str(action.get("value", ""))
        elif isinstance(action, bool):
            mode = "checked" if action else "unchecked"
        elif isinstance(action, str):
            normalized = action.strip()
            lowered = normalized.lower()
            if lowered in self._CHECKBOX_TRUE_VALUES or normalized.startswith("/"):
                mode = "checked"
            elif lowered in self._CHECKBOX_FALSE_VALUES:
                mode = "unchecked"
            else:
                mode = "literal"
                literal_value = normalized
        elif action is not None:
            mode = "literal"
            literal_value = str(action)

        index = self._mode_combo.findData(mode)
        if index < 0:
            index = 0
        self._mode_combo.setCurrentIndex(index)

        if mode == "literal":
            self._text_edit.setText(literal_value)
        else:
            self._text_edit.clear()

        if mode == "column":
            self._column_combo.blockSignals(True)
            if column:
                combo_index = self._column_combo.findText(column)
                if combo_index >= 0:
                    self._column_combo.setCurrentIndex(combo_index)
                else:
                    self._column_combo.insertItem(0, column)
                    self._column_combo.setCurrentIndex(0)
            else:
                if self._column_combo.count():
                    self._column_combo.setCurrentIndex(0)
            self._column_combo.blockSignals(False)
            self._column_fallback.setText(fallback)
        else:
            if self._column_combo.count():
                self._column_combo.setCurrentIndex(0)
            self._column_fallback.clear()

        self._block_updates = False
        self._on_mode_changed(self._mode_combo.currentIndex())

    def clear(self) -> None:
        self.load_action(None)

    def action_spec(self) -> Dict[str, Any] | None:
        mode = self._mode_combo.currentData()
        if mode == "ignore":
            return None
        if mode == "checked":
            return {"mode": "checkbox", "checked": True}
        if mode == "unchecked":
            return {"mode": "checkbox", "checked": False}
        if mode == "literal":
            return {"mode": "literal", "value": self._text_edit.text()}
        if mode == "column":
            column = self._column_combo.currentText().strip()
            data: Dict[str, Any] = {"mode": "column"}
            if column:
                data["column"] = column
            fallback = self._column_fallback.text()
            if fallback:
                data["fallback"] = fallback
            return data
        return None

    def _on_mode_changed(self, index: int) -> None:
        mode = self._mode_combo.itemData(index)
        if mode == "literal":
            self._stack.setCurrentIndex(1)
        elif mode == "column":
            self._stack.setCurrentIndex(2)
        else:
            self._stack.setCurrentIndex(0)
        if not self._block_updates:
            self._emit_changed()

    def _emit_changed(self) -> None:
        if not self._block_updates:
            self.changed.emit()


class _ChoiceCaseEditor(QtWidgets.QWidget):
    """Composite editor for a single conditional case."""

    changed = QtCore.Signal()

    def __init__(
        self,
        columns: Sequence[str],
        targets: Sequence[str],
        *,
        include_match_field: bool,
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._include_match_field = include_match_field
        self._columns = list(columns)
        self._targets = list(targets)
        self._block_updates = False

        layout = QtWidgets.QFormLayout(self)
        layout.setFieldGrowthPolicy(QtWidgets.QFormLayout.ExpandingFieldsGrow)
        layout.setVerticalSpacing(6)
        layout.setHorizontalSpacing(12)
        layout.setLabelAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)

        self._match_edit: QtWidgets.QLineEdit | None = None
        if include_match_field:
            self._match_edit = QtWidgets.QLineEdit()
            self._match_edit.setPlaceholderText("Value to match")
            self._match_edit.textEdited.connect(self._emit_changed)
            layout.addRow("Match value:", self._match_edit)

        self._targets_group = QtWidgets.QGroupBox("Field actions")
        targets_layout = QtWidgets.QFormLayout(self._targets_group)
        targets_layout.setFieldGrowthPolicy(QtWidgets.QFormLayout.ExpandingFieldsGrow)
        targets_layout.setVerticalSpacing(4)
        targets_layout.setHorizontalSpacing(10)
        targets_layout.setContentsMargins(8, 6, 8, 6)
        targets_layout.setLabelAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)

        self._target_editors: Dict[str, _ChoiceTargetActionEditor] = {}
        for target in self._targets:
            editor = _ChoiceTargetActionEditor(target, self._columns)
            editor.changed.connect(self._emit_changed)
            targets_layout.addRow(f"{target}:", editor)
            self._target_editors[target] = editor

        layout.addRow(self._targets_group)

    def set_columns(self, columns: Sequence[str]) -> None:
        self._columns = list(columns)
        for editor in self._target_editors.values():
            editor.set_columns(self._columns)

    def set_targets(self, targets: Sequence[str]) -> None:
        stored_actions = self.actions()
        self._targets = list(targets)
        layout = self._targets_group.layout()
        assert isinstance(layout, QtWidgets.QFormLayout)
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self._target_editors = {}
        for target in self._targets:
            editor = _ChoiceTargetActionEditor(target, self._columns)
            editor.changed.connect(self._emit_changed)
            layout.addRow(f"{target}:", editor)
            if target in stored_actions:
                editor.load_action(stored_actions[target])
            self._target_editors[target] = editor

    def load_case(self, case: Mapping[str, Any] | None) -> None:
        self._block_updates = True
        if self._match_edit is not None:
            value = ""
            if case is not None:
                value = str(case.get("match", ""))
            self._match_edit.setText(value)
        outputs: Mapping[str, Any] = {}
        if case is not None:
            payload = case.get("outputs", {})
            if isinstance(payload, Mapping):
                outputs = payload
        for target, editor in self._target_editors.items():
            editor.load_action(outputs.get(target))
        self._block_updates = False

    def actions(self) -> Dict[str, Any]:
        result: Dict[str, Any] = {}
        for target, editor in self._target_editors.items():
            spec = editor.action_spec()
            if spec is not None:
                result[target] = spec
        return result

    def case_data(self) -> Dict[str, Any]:
        payload: Dict[str, Any] = {"outputs": self.actions()}
        if self._match_edit is not None:
            payload["match"] = self._match_edit.text().strip()
        return payload

    def clear(self) -> None:
        self.load_case(None)

    def _emit_changed(self) -> None:
        if not self._block_updates:
            self.changed.emit()


class _ChoiceConfigWidget(QtWidgets.QWidget):
    """Configuration panel for conditional mappings."""

    def __init__(self, columns: Sequence[str], targets: Sequence[str], parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._all_columns = list(columns)
        self._targets = list(targets)
        self._cases: list[Dict[str, Any]] = []
        self._current_case_index = -1
        self._loading_case = False

        self._source_combo = QtWidgets.QComboBox()
        self._source_combo.addItems(self._all_columns)

        instructions = QtWidgets.QLabel(
            "Define how each value from the data column should toggle checkboxes or fill text fields."
        )
        instructions.setObjectName("choiceInstructions")
        instructions.setWordWrap(True)

        self._cases_list = QtWidgets.QListWidget()
        self._cases_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self._cases_list.setUniformItemSizes(True)
        self._cases_list.currentRowChanged.connect(self._on_case_selected)

        self._case_editor = _ChoiceCaseEditor(self._all_columns, self._targets, include_match_field=True)
        self._case_editor.changed.connect(self._on_case_changed)

        cases_panel = QtWidgets.QWidget()
        cases_layout = QtWidgets.QHBoxLayout(cases_panel)
        cases_layout.setContentsMargins(0, 0, 0, 0)
        cases_layout.setSpacing(8)
        cases_layout.addWidget(self._cases_list, 2)
        cases_layout.addWidget(self._case_editor, 5)
        cases_panel.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.MinimumExpanding)

        self._add_case_button = QtWidgets.QPushButton("Add value")
        self._remove_case_button = QtWidgets.QPushButton("Remove value")
        self._add_case_button.clicked.connect(self._handle_add_case)
        self._remove_case_button.clicked.connect(self._handle_remove_case)

        buttons_row = QtWidgets.QHBoxLayout()
        buttons_row.setSpacing(6)
        buttons_row.addWidget(self._add_case_button)
        buttons_row.addWidget(self._remove_case_button)
        buttons_row.addStretch(1)

        self._fallback_editor = _ChoiceCaseEditor(
            self._all_columns,
            self._targets,
            include_match_field=False,
        )
        self._fallback_editor.changed.connect(self._on_case_changed)

        fallback_group = QtWidgets.QGroupBox("When no value matches")
        fallback_group.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        fallback_layout = QtWidgets.QVBoxLayout(fallback_group)
        fallback_layout.setContentsMargins(8, 8, 8, 8)
        fallback_layout.setSpacing(6)
        fallback_hint = QtWidgets.QLabel("Optional actions to apply when no case is triggered.")
        fallback_hint.setWordWrap(True)
        fallback_layout.addWidget(fallback_hint)
        fallback_layout.addWidget(self._fallback_editor)

        scroll_content = QtWidgets.QWidget()
        scroll_layout = QtWidgets.QVBoxLayout(scroll_content)
        scroll_layout.setContentsMargins(0, 0, 0, 0)
        scroll_layout.setSpacing(10)
        scroll_layout.addWidget(instructions)
        scroll_layout.addWidget(cases_panel, 1)
        scroll_layout.addLayout(buttons_row)
        scroll_layout.addWidget(fallback_group)
        scroll_layout.addStretch(1)

        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setObjectName("choiceScrollArea")
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QtWidgets.QFrame.NoFrame)
        scroll_area.setWidget(scroll_content)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        source_label = QtWidgets.QLabel("Source column:")
        layout.addWidget(source_label)
        layout.addWidget(self._source_combo)
        layout.addWidget(scroll_area)

        self._ensure_case_exists()

    def set_targets(self, targets: Sequence[str]) -> None:
        self._sync_current_case()
        self._targets = list(targets)
        for case in self._cases:
            outputs = case.get("outputs", {})
            if isinstance(outputs, Mapping):
                case["outputs"] = {key: outputs[key] for key in outputs if key in self._targets}
            else:
                case["outputs"] = {}
        fallback_outputs = self._fallback_editor.actions()
        self._case_editor.set_targets(self._targets)
        self._fallback_editor.set_targets(self._targets)
        self._fallback_editor.load_case({"outputs": fallback_outputs})
        self._refresh_case_list()
        if 0 <= self._current_case_index < len(self._cases):
            self._loading_case = True
            self._case_editor.load_case(self._cases[self._current_case_index])
            self._loading_case = False
        elif self._cases:
            self._cases_list.setCurrentRow(0)

    def set_columns(self, columns: Sequence[str]) -> None:
        current = self._source_combo.currentText()
        self._source_combo.blockSignals(True)
        self._source_combo.clear()
        self._source_combo.addItems(columns)
        index = self._source_combo.findText(current)
        if index >= 0:
            self._source_combo.setCurrentIndex(index)
        self._source_combo.blockSignals(False)
        self._all_columns = list(columns)
        self._case_editor.set_columns(self._all_columns)
        self._fallback_editor.set_columns(self._all_columns)

    def load_options(self, options: Dict[str, object], *, default_source: str | None = None) -> None:
        source_value = options.get("source")
        source = source_value if isinstance(source_value, str) else ""
        if not source and default_source:
            source = default_source
        if source:
            index = self._source_combo.findText(source)
            if index >= 0:
                self._source_combo.setCurrentIndex(index)
            else:
                self._source_combo.insertItem(0, source)
                self._source_combo.setCurrentIndex(0)
        elif self._source_combo.count() and self._source_combo.currentIndex() < 0:
            self._source_combo.setCurrentIndex(0)

        cases_payload = options.get("cases")
        parsed_cases: list[Dict[str, Any]] = []
        if isinstance(cases_payload, Mapping):
            for match_value, outputs in cases_payload.items():
                parsed_cases.append(
                    {
                        "match": str(match_value),
                        "outputs": self._normalize_outputs(outputs),
                    }
                )
        elif isinstance(cases_payload, list):
            for case in cases_payload:
                if isinstance(case, Mapping):
                    parsed_cases.append(
                        {
                            "match": str(case.get("match", "")),
                            "outputs": self._normalize_outputs(case.get("outputs", {})),
                        }
                    )
        elif not cases_payload:
            case_map = options.get("case_map")
            if isinstance(case_map, Mapping):
                for match_value, outputs in case_map.items():
                    parsed_cases.append(
                        {
                            "match": str(match_value),
                            "outputs": self._normalize_outputs(outputs),
                        }
                    )
        case_map = options.get("case_map")
        if isinstance(case_map, Mapping) and case_map:
            existing = [str(case.get("match", "")).strip() for case in parsed_cases]
            blanks = [case for case in parsed_cases if not str(case.get("match", "")).strip()]
            for match_value, outputs in case_map.items():
                key = str(match_value).strip()
                if not key:
                    continue
                if key in existing:
                    continue
                if blanks:
                    slot = blanks.pop(0)
                    slot["match"] = key
                    slot["outputs"] = self._normalize_outputs(outputs)
                    existing.append(key)
                else:
                    parsed_cases.append(
                        {
                            "match": key,
                            "outputs": self._normalize_outputs(outputs),
                        }
                    )
                    existing.append(key)
        self._cases = parsed_cases or []
        self._refresh_case_list()
        if self._cases:
            self._cases_list.setCurrentRow(0)
        else:
            self._ensure_case_exists()

        default_payload = options.get("default", {})
        if isinstance(default_payload, Mapping):
            self._fallback_editor.load_case({"outputs": self._normalize_outputs(default_payload)})
        else:
            self._fallback_editor.clear()

    def build_options(self) -> Dict[str, object]:
        self._sync_current_case()
        cases_map: Dict[str, Dict[str, Any]] = {}
        cases_list: list[Dict[str, Any]] = []
        for case in self._cases:
            match_value = str(case.get("match", "")).strip()
            if not match_value:
                continue
            outputs: Dict[str, Any] = {}
            for target, action in (case.get("outputs") or {}).items():
                simplified = self._simplify_action(action)
                if simplified is not None:
                    outputs[str(target)] = simplified
            cases_list.append({"match": match_value, "outputs": outputs})
            cases_map[match_value] = outputs

        default_outputs: Dict[str, Any] = {}
        for target, action in self._fallback_editor.actions().items():
            simplified = self._simplify_action(action)
            if simplified is not None:
                default_outputs[target] = simplified

        payload: Dict[str, object] = {
            "source": self._source_combo.currentText(),
        }
        if cases_list:
            payload["cases"] = cases_list
            payload["case_map"] = cases_map
        else:
            payload["cases"] = {}
        if default_outputs:
            payload["default"] = default_outputs
        return payload

    def ensure_source_selected(self, preferred: str | None = None) -> None:
        if self._source_combo.currentIndex() >= 0 and self._source_combo.currentText():
            return
        candidate = preferred or ""
        if candidate:
            index = self._source_combo.findText(candidate)
            if index >= 0:
                self._source_combo.setCurrentIndex(index)
                return
        if self._source_combo.count():
            self._source_combo.setCurrentIndex(0)

    def validate(self) -> tuple[bool, str | None]:
        self._sync_current_case()
        for case in self._cases:
            match_value = str(case.get("match", "")).strip()
            if not match_value:
                return False, "Enter a match value for each conditional choice."
            outputs = case.get("outputs", {})
            if not outputs:
                return False, f"Add at least one field action for '{match_value}'."
            for target, action in outputs.items():
                if isinstance(action, Mapping):
                    mode = str(action.get("mode") or action.get("kind") or action.get("type") or "").lower()
                    if not mode and "column" in action:
                        mode = "column"
                    if mode == "column" and not str(action.get("column", "")).strip():
                        return False, f"Select a column for '{target}' when '{match_value}' is matched."
        for target, action in self._fallback_editor.actions().items():
            if isinstance(action, Mapping):
                mode = str(action.get("mode") or action.get("kind") or action.get("type") or "").lower()
                if not mode and "column" in action:
                    mode = "column"
                if mode == "column" and not str(action.get("column", "")).strip():
                    return False, f"Select a column for '{target}' in the fallback configuration."
        return True, None

    def _ensure_case_exists(self) -> None:
        if not self._cases:
            self._cases.append({"match": "", "outputs": {}})
            self._refresh_case_list()
            self._cases_list.setCurrentRow(0)

    def _refresh_case_list(self) -> None:
        self._cases_list.blockSignals(True)
        self._cases_list.clear()
        for case in self._cases:
            self._cases_list.addItem(self._format_case_label(case))
        self._cases_list.blockSignals(False)
        if 0 <= self._current_case_index < len(self._cases):
            self._cases_list.setCurrentRow(self._current_case_index)
        else:
            self._current_case_index = self._cases_list.currentRow()

    def _format_case_label(self, case: Mapping[str, Any]) -> str:
        match_value = str(case.get("match", "")).strip()
        if match_value:
            return match_value
        outputs = case.get("outputs", {})
        if outputs:
            return ", ".join(str(target) for target in outputs.keys())
        return "New value"

    def _on_case_selected(self, index: int) -> None:
        if self._loading_case:
            return
        if 0 <= self._current_case_index < len(self._cases):
            self._cases[self._current_case_index] = self._case_editor.case_data()
            self._update_case_label(self._current_case_index)
        self._current_case_index = index
        self._loading_case = True
        if 0 <= index < len(self._cases):
            self._case_editor.load_case(self._cases[index])
        else:
            self._case_editor.clear()
        self._loading_case = False

    def _on_case_changed(self) -> None:
        if self._loading_case or not (0 <= self._current_case_index < len(self._cases)):
            return
        self._cases[self._current_case_index] = self._case_editor.case_data()
        self._update_case_label(self._current_case_index)

    def _update_case_label(self, index: int) -> None:
        if 0 <= index < self._cases_list.count():
            self._cases_list.item(index).setText(self._format_case_label(self._cases[index]))

    def _sync_current_case(self) -> None:
        self._on_case_changed()

    def _handle_add_case(self) -> None:
        self._sync_current_case()
        self._cases.append({"match": "", "outputs": {}})
        self._refresh_case_list()
        self._cases_list.setCurrentRow(len(self._cases) - 1)

    def _handle_remove_case(self) -> None:
        index = self._cases_list.currentRow()
        if index < 0:
            return
        self._cases.pop(index)
        if not self._cases:
            self._cases.append({"match": "", "outputs": {}})
        self._refresh_case_list()
        new_index = min(index, len(self._cases) - 1)
        self._cases_list.setCurrentRow(new_index)

    def _simplify_action(self, action: Any) -> Any:
        if not isinstance(action, Mapping):
            return action
        mode = str(action.get("mode") or action.get("kind") or action.get("type") or "").lower()
        if not mode and "column" in action:
            mode = "column"
        if not mode and "value" in action:
            mode = "literal"
        if mode == "checkbox":
            checked = action.get("checked")
            if isinstance(checked, bool):
                on_value = action.get("value") or action.get("checked_value") or action.get("on")
                off_value = action.get("unchecked_value") or action.get("off")
                if on_value is None and off_value is None:
                    return checked
                return on_value if checked else off_value
            return bool(checked)
        if mode in {"literal", "text", "value"}:
            return action.get("value", "")
        if mode == "column":
            result: Dict[str, Any] = {"mode": "column"}
            column = action.get("column")
            if column:
                result["column"] = column
            fallback = action.get("fallback")
            if fallback:
                result["fallback"] = fallback
            format_pattern = action.get("format")
            if format_pattern:
                result["format"] = format_pattern
            return result
        if mode == "raw":
            return action.get("value")
        return action

    def _normalize_outputs(self, payload: Any) -> Dict[str, Any]:
        result: Dict[str, Any] = {}
        if isinstance(payload, Mapping):
            for key, value in payload.items():
                spec = self._normalize_action(value)
                if spec:
                    result[str(key)] = spec
        return result

    def _normalize_action(self, value: Any) -> Dict[str, Any]:
        if isinstance(value, Mapping):
            data = dict(value)
            mode = str(data.get("mode") or data.get("kind") or data.get("type") or "").lower()
            if not mode:
                if "column" in data:
                    mode = "column"
                elif "checked" in data:
                    mode = "checkbox"
                elif "value" in data:
                    mode = "literal"
            result: Dict[str, Any] = {"mode": mode} if mode else {}
            for key in ("column", "fallback", "format", "value", "checked", "checked_value", "unchecked_value", "on", "off"):
                if key in data:
                    result[key] = data[key]
            return result
        if isinstance(value, bool):
            return {"mode": "checkbox", "checked": value}
        if isinstance(value, str):
            normalized = value.strip()
            lowered = normalized.lower()
            if lowered in _ChoiceTargetActionEditor._CHECKBOX_TRUE_VALUES or normalized.startswith("/"):
                return {"mode": "checkbox", "checked": True}
            if lowered in _ChoiceTargetActionEditor._CHECKBOX_FALSE_VALUES:
                return {"mode": "checkbox", "checked": False}
            return {"mode": "literal", "value": value}
        if value is not None:
            return {"mode": "literal", "value": value}
        return {}
class RuleEditorDialog(QtWidgets.QDialog):
    """Modal dialog that edits a mapping rule."""

    def __init__(
        self,
        field_name: str,
        rule: MappingRule | None,
        available_fields: Sequence[str],
        available_columns: Sequence[str],
        parent: QtWidgets.QWidget | None = None,
        *,
        remove_callback: Optional[Callable[[], None]] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Edit Mapping Rule - {field_name}")
        self.resize(920, 560)

        self._field_name = field_name
        self._available_fields = list(available_fields)
        self._available_columns = list(available_columns)
        self._remove_callback = remove_callback

        self._types_combo = QtWidgets.QComboBox()
        self._types_combo.addItem("Direct value", RuleType.VALUE)
        self._types_combo.addItem("Literal", RuleType.LITERAL)
        self._types_combo.addItem("Conditional (choices)", RuleType.CHOICE)
        self._types_combo.addItem("Concatenate columns", RuleType.CONCAT)

        self._targets_list = _TargetsSelector()
        self._targets_list.setMinimumHeight(120)
        self._targets_list.setMinimumWidth(220)
        self._targets_list.selectionChanged.connect(self._on_targets_changed)

        self._value_widget = _ValueConfigWidget(self._available_columns)
        self._literal_widget = _LiteralConfigWidget()
        self._concat_widget = _ConcatConfigWidget(self._available_columns)
        self._choice_widget = _ChoiceConfigWidget(self._available_columns, [])

        info_icon = self.style().standardIcon(QtWidgets.QStyle.SP_MessageBoxInformation)
        self._help_icon = QtWidgets.QLabel()
        pixmap = info_icon.pixmap(20, 20)
        self._help_icon.setPixmap(pixmap)
        palette = self.palette()
        highlight = palette.color(QtGui.QPalette.Highlight)
        info_background = QtGui.QColor(highlight)
        info_background.setAlpha(30)

        self._help_icon.setAutoFillBackground(True)
        icon_palette = self._help_icon.palette()
        icon_palette.setColor(QtGui.QPalette.Window, info_background)
        self._help_icon.setPalette(icon_palette)

        frame = QtWidgets.QFrame()
        frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        frame.setAutoFillBackground(True)
        frame_palette = frame.palette()
        frame_palette.setColor(QtGui.QPalette.Window, info_background)
        frame_palette.setColor(QtGui.QPalette.WindowText, palette.color(QtGui.QPalette.WindowText))
        frame.setPalette(frame_palette)
        frame.setStyleSheet("QFrame { border: 1px solid palette(highlight); border-radius: 6px; }")

        self._help_label = QtWidgets.QLabel()
        self._help_label.setWordWrap(True)
        self._help_label.setObjectName("ruleHelpLabel")
        self._help_label.setAlignment(QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft)

        self._stack = _AutoSizingStack()
        self._stack.addWidget(self._value_widget)
        self._stack.addWidget(self._literal_widget)
        self._stack.addWidget(self._choice_widget)
        self._stack.addWidget(self._concat_widget)

        self._types_combo.currentIndexChanged.connect(self._on_rule_type_changed)

        targets_group = QtWidgets.QGroupBox("PDF fields to populate")
        targets_layout = QtWidgets.QVBoxLayout(targets_group)
        targets_layout.addWidget(self._targets_list)
        targets_group.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)

        type_group = QtWidgets.QGroupBox("Rule configuration")
        type_group.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        type_layout = QtWidgets.QVBoxLayout(type_group)
        type_layout.addWidget(self._types_combo)
        help_row = QtWidgets.QHBoxLayout()
        help_row.setContentsMargins(0, 0, 0, 0)
        help_row.setSpacing(8)
        help_row.addWidget(self._help_icon, 0, QtCore.Qt.AlignTop)
        help_row.addWidget(self._help_label, 1)
        type_layout.addLayout(help_row)
        type_layout.addWidget(self._stack, 1)

        split_panel = QtWidgets.QWidget()
        split_layout = QtWidgets.QHBoxLayout(split_panel)
        split_layout.setContentsMargins(0, 0, 0, 0)
        split_layout.setSpacing(16)
        split_layout.addWidget(targets_group, 1)
        split_layout.addWidget(type_group, 2)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(16)
        layout.addWidget(split_panel, 1)
        layout.addWidget(self._build_buttons(include_remove=remove_callback is not None))

        initial_rule = rule or MappingRule.from_direct_column(
            field_name,
            self._available_columns[0] if self._available_columns else "",
        )
        self._load_rule(initial_rule)

    def _build_buttons(self, *, include_remove: bool) -> QtWidgets.QWidget:
        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        container = QtWidgets.QWidget()
        layout = QtWidgets.QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        if include_remove:
            remove_button = QtWidgets.QPushButton("Remove Mapping")
            remove_button.setObjectName("removeMappingButton")
            remove_button.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_TrashIcon))
            remove_button.setAutoDefault(False)
            remove_button.clicked.connect(self._handle_remove_clicked)
            layout.addWidget(remove_button)
            layout.addStretch(1)
        else:
            layout.addStretch(1)
        layout.addWidget(button_box)
        return container

    def _handle_remove_clicked(self) -> None:
        if self._remove_callback is None:
            return
        response = QtWidgets.QMessageBox.question(
            self,
            "Remove Mapping",
            f"Remove mapping for '{self._field_name}'?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )
        if response != QtWidgets.QMessageBox.Yes:
            return
        self._remove_callback()
        self.reject()

    def _load_rule(self, rule: MappingRule) -> None:
        selected_targets = rule.targets or [rule.name]
        self._targets_list.set_targets(self._available_fields, selected_targets)

        rule_type = rule.type_enum()
        index = self._types_combo.findData(rule_type)
        if index >= 0:
            self._types_combo.setCurrentIndex(index)

        self._value_widget.set_columns(self._available_columns)
        self._value_widget.load_options(rule.options)
        self._literal_widget.load_options(rule.options)
        self._concat_widget.set_columns(self._available_columns)
        self._concat_widget.load_options(rule.options)
        self._choice_widget.set_targets(selected_targets)
        self._choice_widget.set_columns(self._available_columns)
        preferred_source = ""
        source_option = rule.options.get("source")
        if isinstance(source_option, str):
            preferred_source = source_option
        else:
            column_option = rule.options.get("column")
            if isinstance(column_option, str):
                preferred_source = column_option
        if not preferred_source:
            preferred_source = self._value_widget.current_column()
        self._choice_widget.load_options(rule.options, default_source=preferred_source)
        self._choice_widget.ensure_source_selected(preferred_source)
        self._update_help(rule_type)

    def _on_targets_changed(self) -> None:
        targets = self._targets_list.selected_targets()
        if not targets:
            return
        self._choice_widget.set_targets(targets)

    def _on_rule_type_changed(self, index: int) -> None:
        data = self._types_combo.itemData(index)
        rule_type = self._coerce_rule_type(data)
        if rule_type is RuleType.VALUE:
            self._stack.setCurrentWidget(self._value_widget)
        elif rule_type is RuleType.LITERAL:
            self._stack.setCurrentWidget(self._literal_widget)
        elif rule_type is RuleType.CHOICE:
            self._stack.setCurrentWidget(self._choice_widget)
            self._choice_widget.set_targets(self._targets_list.selected_targets() or [self._field_name])
            self._choice_widget.ensure_source_selected(self._value_widget.current_column())
        elif rule_type is RuleType.CONCAT:
            self._stack.setCurrentWidget(self._concat_widget)
        self._update_help(rule_type)

    def selected_rule(self) -> MappingRule:
        data = self._types_combo.currentData()
        rule_type = self._coerce_rule_type(data)
        targets = self._targets_list.selected_targets() or [self._field_name]
        options = self._gather_options(rule_type, targets)
        return MappingRule(
            name=self._field_name,
            rule_type=rule_type,
            targets=targets,
            options=options,
        )

    def _gather_options(self, rule_type: RuleType, targets: Sequence[str]) -> Dict[str, object]:
        rule_type = self._coerce_rule_type(rule_type)
        if rule_type is RuleType.VALUE:
            return self._value_widget.build_options()
        if rule_type is RuleType.LITERAL:
            return self._literal_widget.build_options()
        if rule_type is RuleType.CONCAT:
            return self._concat_widget.build_options()
        if rule_type is RuleType.CHOICE:
            self._choice_widget.set_targets(targets)
            return self._choice_widget.build_options()
        return {}

    def accept(self) -> None:
        targets = self._targets_list.selected_targets()
        if not targets:
            QtWidgets.QMessageBox.warning(
                self,
                "Validate Rule",
                "Select at least one PDF field to populate.",
            )
            return
        data = self._types_combo.currentData()
        rule_type = self._coerce_rule_type(data)
        options = self._gather_options(rule_type, targets)
        if not self._validate(rule_type, options):
            return
        super().accept()

    def _validate(self, rule_type: RuleType, options: Dict[str, object]) -> bool:
        rule_type = self._coerce_rule_type(rule_type)
        if rule_type is RuleType.VALUE:
            column = str(options.get("column", "")).strip()
            if not column:
                QtWidgets.QMessageBox.warning(self, "Validate Rule", "Select a source column.")
                return False
        if rule_type is RuleType.CHOICE:
            source = str(options.get("source", "")).strip()
            if not source:
                QtWidgets.QMessageBox.warning(self, "Validate Rule", "Select the source column for the conditional rule.")
                return False
            valid, message = self._choice_widget.validate()
            if not valid:
                QtWidgets.QMessageBox.warning(
                    self,
                    "Validate Rule",
                    message or "Configure at least one conditional value.",
                )
                return False
        if rule_type is RuleType.CONCAT:
            columns = options.get("columns", [])
            if not columns:
                QtWidgets.QMessageBox.warning(self, "Validate Rule", "Select at least one column to concatenate.")
                return False
        return True

    def _update_help(self, rule_type: RuleType) -> None:
        text = RULE_HELP_TEXT.get(rule_type, "")
        self._help_label.setText(text)

    @staticmethod
    def _coerce_rule_type(data: object) -> RuleType:
        if isinstance(data, RuleType):
            return data
        if isinstance(data, str):
            return RuleType(data)
        if data is None:
            return RuleType.VALUE
        return RuleType(data)
