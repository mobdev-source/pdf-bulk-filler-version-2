"""Dialog for configuring mapping rules via the GUI."""

from __future__ import annotations

from typing import Callable, Dict, Iterable, List, Optional, Sequence

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
        column = options.get("column", "")
        index = self._column_combo.findText(column)
        if index >= 0:
            self._column_combo.setCurrentIndex(index)
        elif self._column_combo.count():
            self._column_combo.setCurrentIndex(0)
        self._default_edit.setText(options.get("default", ""))
        self._format_edit.setText(options.get("format", ""))

    def build_options(self) -> Dict[str, str]:
        return {
            "column": self._column_combo.currentText(),
            "default": self._default_edit.text(),
            "format": self._format_edit.text(),
        }


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
        selected_set = {str(name) for name in selected or []}
        self._column_list.clear()
        for column in self._columns:
            item = QtWidgets.QListWidgetItem(column)
            item.setFlags(
                QtCore.Qt.ItemIsEnabled
                | QtCore.Qt.ItemIsSelectable
                | QtCore.Qt.ItemIsUserCheckable
                | QtCore.Qt.ItemIsDragEnabled
            )
            state = QtCore.Qt.Checked if column in selected_set else QtCore.Qt.Unchecked
            item.setCheckState(state)
            self._column_list.addItem(item)

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


class _ChoiceConfigWidget(QtWidgets.QWidget):
    """Configuration panel for conditional mappings."""

    def __init__(self, columns: Sequence[str], targets: Sequence[str], parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._all_columns = list(columns)
        self._targets: list[str] = list(targets)
        self._default_fields: Dict[str, QtWidgets.QLineEdit] = {}

        self._source_combo = QtWidgets.QComboBox()
        self._source_combo.addItems(self._all_columns)

        self._cases_table = QtWidgets.QTableWidget(0, 1)
        self._cases_table.setHorizontalHeaderLabels(["Match Value"])
        self._cases_table.horizontalHeader().setStretchLastSection(True)
        self._cases_table.setEditTriggers(QtWidgets.QAbstractItemView.AllEditTriggers)

        self._add_case_button = QtWidgets.QPushButton("Add Case")
        self._remove_case_button = QtWidgets.QPushButton("Remove Selected")

        self._add_case_button.clicked.connect(self._add_case)
        self._remove_case_button.clicked.connect(self._remove_case)

        button_row = QtWidgets.QHBoxLayout()
        button_row.addWidget(self._add_case_button)
        button_row.addWidget(self._remove_case_button)
        button_row.addStretch(1)

        instructions = QtWidgets.QLabel(
            "Add a row for each possible value in the data column, then specify what each PDF field "
            "should display for that value."
        )
        instructions.setWordWrap(True)

        self._default_layout = QtWidgets.QFormLayout()

        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(QtWidgets.QLabel("Source column:"))
        layout.addWidget(self._source_combo)
        layout.addWidget(instructions)
        layout.addLayout(button_row)
        layout.addWidget(self._cases_table, 4)
        defaults_group = QtWidgets.QGroupBox("Default outputs")
        defaults_group.setLayout(self._default_layout)
        layout.addWidget(defaults_group)

        self.set_targets(targets)
        self.set_columns(columns)

    def set_targets(self, targets: Sequence[str]) -> None:
        existing_cases = self._collect_cases()
        existing_defaults = {name: field.text() for name, field in self._default_fields.items()}
        self._targets = list(targets)
        column_labels = ["Match Value"] + list(self._targets)
        self._cases_table.setColumnCount(len(column_labels))
        self._cases_table.setHorizontalHeaderLabels(column_labels)
        self._cases_table.horizontalHeader().setStretchLastSection(True)

        self._cases_table.setRowCount(len(existing_cases))
        for row_index, case in enumerate(existing_cases):
            match_value = case.get("match", "")
            self._cases_table.setItem(row_index, 0, QtWidgets.QTableWidgetItem(match_value))
            outputs = case.get("outputs", {})
            for column_offset, target in enumerate(self._targets, start=1):
                value = outputs.get(target, "")
                self._cases_table.setItem(row_index, column_offset, QtWidgets.QTableWidgetItem(value))

        for i in reversed(range(self._default_layout.count())):
            item = self._default_layout.takeAt(i)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self._default_fields = {}
        for target in self._targets:
            edit = QtWidgets.QLineEdit(existing_defaults.get(target, ""))
            self._default_layout.addRow(f"{target}:", edit)
            self._default_fields[target] = edit

    def set_columns(self, columns: Sequence[str]) -> None:
        current = self._source_combo.currentText()
        self._source_combo.blockSignals(True)
        self._source_combo.clear()
        self._source_combo.addItems(columns)
        index = self._source_combo.findText(current)
        if index >= 0:
            self._source_combo.setCurrentIndex(index)
        self._source_combo.blockSignals(False)

    def load_options(self, options: Dict[str, object]) -> None:
        source = str(options.get("source", ""))
        index = self._source_combo.findText(source)
        if index >= 0:
            self._source_combo.setCurrentIndex(index)
        cases = options.get("cases", {})
        if isinstance(cases, dict):
            parsed_cases = []
            for match_value, outputs in cases.items():
                row_outputs = {}
                if isinstance(outputs, dict):
                    for target in self._targets:
                        row_outputs[target] = str(outputs.get(target, ""))
                else:
                    row_outputs = {target: str(outputs) for target in self._targets}
                parsed_cases.append({"match": str(match_value), "outputs": row_outputs})
            self._cases_table.setRowCount(len(parsed_cases))
            for row_index, case in enumerate(parsed_cases):
                self._cases_table.setItem(row_index, 0, QtWidgets.QTableWidgetItem(case["match"]))
                for column_offset, target in enumerate(self._targets, start=1):
                    value = case["outputs"].get(target, "")
                    self._cases_table.setItem(row_index, column_offset, QtWidgets.QTableWidgetItem(value))
        defaults = options.get("default", {})
        if isinstance(defaults, dict):
            for target, edit in self._default_fields.items():
                edit.setText(str(defaults.get(target, "")))

    def build_options(self) -> Dict[str, object]:
        cases = {}
        for case in self._collect_cases():
            match_value = case.get("match", "")
            if not match_value:
                continue
            cases[match_value] = {
                target: case["outputs"].get(target, "")
                for target in self._targets
            }
        defaults = {target: edit.text() for target, edit in self._default_fields.items()}
        return {
            "source": self._source_combo.currentText(),
            "cases": cases,
            "default": defaults,
        }

    def _collect_cases(self) -> List[Dict[str, object]]:
        data: List[Dict[str, object]] = []
        for row in range(self._cases_table.rowCount()):
            match_item = self._cases_table.item(row, 0)
            match_value = match_item.text() if match_item else ""
            outputs: Dict[str, str] = {}
            for column_offset, target in enumerate(self._targets, start=1):
                cell = self._cases_table.item(row, column_offset)
                outputs[target] = cell.text() if cell else ""
            data.append({"match": match_value, "outputs": outputs})
        return data

    def _add_case(self) -> None:
        row = self._cases_table.rowCount()
        self._cases_table.insertRow(row)
        self._cases_table.setItem(row, 0, QtWidgets.QTableWidgetItem(""))
        for column_offset in range(1, len(self._targets) + 1):
            self._cases_table.setItem(row, column_offset, QtWidgets.QTableWidgetItem(""))

    def _remove_case(self) -> None:
        row = self._cases_table.currentRow()
        if row >= 0:
            self._cases_table.removeRow(row)


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
        self.resize(540, 480)

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

        self._stack = QtWidgets.QStackedWidget()
        self._stack.addWidget(self._value_widget)
        self._stack.addWidget(self._literal_widget)
        self._stack.addWidget(self._choice_widget)
        self._stack.addWidget(self._concat_widget)

        self._types_combo.currentIndexChanged.connect(self._on_rule_type_changed)

        targets_group = QtWidgets.QGroupBox("PDF fields to populate")
        targets_layout = QtWidgets.QVBoxLayout(targets_group)
        targets_layout.addWidget(self._targets_list)

        type_group = QtWidgets.QGroupBox("Rule configuration")
        type_layout = QtWidgets.QVBoxLayout(type_group)
        type_layout.addWidget(self._types_combo)
        help_row = QtWidgets.QHBoxLayout()
        help_row.setContentsMargins(0, 0, 0, 0)
        help_row.setSpacing(8)
        help_row.addWidget(self._help_icon, 0, QtCore.Qt.AlignTop)
        help_row.addWidget(self._help_label, 1)
        type_layout.addLayout(help_row)
        type_layout.addWidget(self._stack, 1)

        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(targets_group)
        layout.addWidget(type_group, 1)
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
        self._choice_widget.load_options(rule.options)
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
