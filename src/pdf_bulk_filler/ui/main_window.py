"""PySide6 main window containing the drag-and-drop mapping interface."""

from __future__ import annotations

from dataclasses import dataclass, field
import re
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, Mapping, Optional, Sequence

import pandas as pd
from PySide6 import QtCore, QtGui, QtWidgets
import fitz
from qfluentwidgets import (
    FluentIcon as FI,
    Theme,
    ThemeColor,
    setTheme,
    setThemeColor,
)
from qfluentwidgets.common.config import qconfig


def get_fluent_icon(*names: str, default: FI = FI.INFO) -> QtGui.QIcon:
    """Return the first available Fluent icon tinted for the active theme."""
    icon_choice = default
    for name in names:
        icon_candidate = getattr(FI, name, None)
        if icon_candidate is not None:
            icon_choice = icon_candidate
            break

    light_color = QtGui.QColor("#1f1f23")
    dark_color = QtGui.QColor("#f5f5f8")
    return icon_choice.colored(light_color, dark_color).qicon()


THEME_SETTINGS_KEY = "ui/themeMode"
DEFAULT_THEME_MODE = "system"
THEME_LABELS = {
    "system": "System (Auto)",
    "light": "Light",
    "dark": "Dark",
}
THEME_MAP = {
    "system": Theme.AUTO,
    "light": Theme.LIGHT,
    "dark": Theme.DARK,
}

ACCENT_COLOR = QtGui.QColor("#3A7BD5")

from pdf_bulk_filler.data.loader import DataLoader, DataSample
from pdf_bulk_filler.mapping.manager import MappingManager, MappingModel
from pdf_bulk_filler.mapping.rules import MappingRule, evaluate_rules
from pdf_bulk_filler.pdf.engine import PdfEngine, PdfField, PdfTemplate
from pdf_bulk_filler.ui.rule_editor import RuleEditorDialog
from pdf_bulk_filler.ui.workers import PdfGenerationWorker

_FILENAME_SANITIZE_PATTERN = re.compile(r"[^A-Za-z0-9._-]+")
_FILENAME_WHITESPACE_PATTERN = re.compile(r"\s+")
_MAX_FILENAME_LENGTH = 120
GENERATION_MODE_KEY = "generation/mode"
GENERATION_DIR_KEY = "generation/directory"
GENERATION_FILE_KEY = "generation/file"
GENERATION_COLUMNS_KEY = "generation/columns"
GENERATION_PREFIX_KEY = "generation/prefix"
GENERATION_SUFFIX_KEY = "generation/suffix"
GENERATION_SEPARATOR_KEY = "generation/separator"
GENERATION_READ_ONLY_KEY = "generation/read_only"


def _sanitize_filename_token(value: object) -> str:
    """Return a filesystem-safe token derived from ``value``."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = _FILENAME_WHITESPACE_PATTERN.sub("_", text)
    text = _FILENAME_SANITIZE_PATTERN.sub("_", text)
    return text.strip("_")[:_MAX_FILENAME_LENGTH]


@dataclass
class GenerationOptions:
    """User selections for PDF generation."""

    mode: str  # "per_entry" or "combined"
    destination_dir: Path | None
    combined_path: Path | None
    columns: list[str]
    prefix: str
    suffix: str
    separator: str
    read_only: bool


class FilenameBuilder:
    """Callable responsible for generating unique filenames per row."""

    def __init__(
        self,
        columns: Sequence[str],
        *,
        prefix: str = "",
        suffix: str = "",
        separator: str = "_",
        index_field: str = "id",
    ) -> None:
        self._columns = list(columns)
        self._prefix = prefix.strip()
        self._suffix = suffix.strip()
        self._separator = separator or "_"
        self._index_field = index_field
        self._counts: Dict[str, int] = {}

    def __call__(self, row: Mapping[str, Any], index: int) -> str:
        base = self._compose_base(row, index)
        normalized = base or f"{index:05d}"
        normalized = normalized[:_MAX_FILENAME_LENGTH]

        count = self._counts.get(normalized, 0)
        self._counts[normalized] = count + 1
        if count:
            suffix = f"{self._separator or '_'}{count+1:02d}"
            available = max(1, _MAX_FILENAME_LENGTH - len(suffix))
            trimmed = normalized[:available].rstrip("_-. ")
            normalized = f"{trimmed}{suffix}"
        return normalized or f"{index:05d}"

    def preview(self, row: Mapping[str, Any], index: int = 1) -> str:
        """Return a sample filename without mutating internal state."""
        temp = FilenameBuilder(
            self._columns,
            prefix=self._prefix,
            suffix=self._suffix,
            separator=self._separator,
            index_field=self._index_field,
        )
        return temp(row, index)

    def _compose_base(self, row: Mapping[str, Any], index: int) -> str:
        tokens: list[str] = []
        if self._prefix:
            tokens.append(_sanitize_filename_token(self._prefix))
        for column in self._columns:
            value = row.get(column)
            token = _sanitize_filename_token(value)
            if token:
                tokens.append(token)
        if self._suffix:
            tokens.append(_sanitize_filename_token(self._suffix))

        if not tokens:
            fallback = row.get(self._index_field) if isinstance(row, Mapping) else None
            token = _sanitize_filename_token(fallback) or f"{index:05d}"
            tokens.append(token)

        separator = self._separator if self._separator is not None else "_"
        joined = separator.join(token for token in tokens if token)
        return joined.strip("_-. ")


class DataFrameModel(QtCore.QAbstractTableModel):
    """A lightweight table model backed by a pandas DataFrame."""

    def __init__(self, frame: Optional[pd.DataFrame] = None) -> None:
        super().__init__()
        self._frame = frame or pd.DataFrame()

    def update(self, frame: pd.DataFrame) -> None:
        self.beginResetModel()
        self._frame = frame
        self.endResetModel()

    def rowCount(self, parent: QtCore.QModelIndex | None = None) -> int:  # noqa: N802
        return 0 if parent and parent.isValid() else len(self._frame.index)

    def columnCount(self, parent: QtCore.QModelIndex | None = None) -> int:  # noqa: N802
        return 0 if parent and parent.isValid() else len(self._frame.columns)

    def data(self, index: QtCore.QModelIndex, role: int = QtCore.Qt.DisplayRole):
        if not index.isValid() or role not in (QtCore.Qt.DisplayRole, QtCore.Qt.EditRole):
            return None
        value = self._frame.iat[index.row(), index.column()]
        if pd.isna(value):
            return ""
        return str(value)

    def headerData(self, section: int, orientation: QtCore.Qt.Orientation, role: int = QtCore.Qt.DisplayRole):  # noqa: N802
        if role != QtCore.Qt.DisplayRole:
            return None
        if orientation == QtCore.Qt.Horizontal:
            try:
                return str(self._frame.columns[section])
            except IndexError:
                return ""
        return str(self._frame.index[section])


class ColumnListWidget(QtWidgets.QListWidget):
    """Displays column names and provides drag support."""

    columnActivated = QtCore.Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self._all_columns: list[str] = []
        self._filter_text: str = ""
        self.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.setDragEnabled(True)
        self.setAlternatingRowColors(True)
        self.itemClicked.connect(self._emit_column_activated)
        self.itemActivated.connect(self._emit_column_activated)

    def set_columns(self, columns: Iterable[str]) -> None:
        self._all_columns = [str(column) for column in columns]
        self._rebuild_visible_items()

    def clear(self) -> None:  # noqa: D401
        """Clear visible and cached columns."""
        super().clear()
        self._all_columns = []
        self._filter_text = ""

    def apply_filter(self, text: str) -> None:
        normalized = (text or "").strip().lower()
        if normalized == self._filter_text:
            return
        self._filter_text = normalized
        self._rebuild_visible_items()

    def mimeData(self, items: list[QtWidgets.QListWidgetItem]) -> QtCore.QMimeData:  # noqa: N802
        mime = QtCore.QMimeData()
        if items:
            mime.setText(items[0].text())
        return mime

    def _rebuild_visible_items(self) -> None:
        selected_text = self.currentItem().text() if self.currentItem() else None
        self.setUpdatesEnabled(False)
        super().clear()
        for column in self._all_columns:
            if self._filter_text and self._filter_text not in column.lower():
                continue
            self.addItem(column)
        self.setUpdatesEnabled(True)
        if selected_text:
            matches = self.findItems(selected_text, QtCore.Qt.MatchExactly)
            if matches:
                self.setCurrentItem(matches[0])

    def _emit_column_activated(self, item: QtWidgets.QListWidgetItem | None) -> None:
        if item is not None:
            self.columnActivated.emit(item.text())


class TooltipStyler(QtCore.QObject):
    """Ensure Qt tooltips mirror the application's theme aesthetic."""

    def __init__(self, parent: QtCore.QObject | None = None) -> None:
        super().__init__(parent)
        self._background = "#ffffff"
        self._text = "#1f1f23"
        self._border = QtGui.QColor(ACCENT_COLOR)
        self._shadow_color = QtGui.QColor(31, 31, 35, 90)
        self._padding = (4, 10)

        app = QtWidgets.QApplication.instance()
        if app is not None:
            app.installEventFilter(self)

    def update_theme(
        self,
        *,
        background: str,
        text: str,
        border: QtGui.QColor,
        shadow: QtGui.QColor,
    ) -> None:
        self._background = background
        self._text = text
        self._border = border
        self._shadow_color = shadow
        self._restyle_active_tooltip()

    def eventFilter(self, obj: QtCore.QObject, event: QtCore.QEvent) -> bool:
        if isinstance(obj, QtWidgets.QWidget) and obj.objectName() in {"qt_tip_label", "qt_tooltip_label"}:
            if event.type() in {
                QtCore.QEvent.Show,
                QtCore.QEvent.PaletteChange,
                QtCore.QEvent.Resize,
            }:
                QtCore.QTimer.singleShot(0, lambda widget=obj: self._apply_styles(widget))
        return super().eventFilter(obj, event)

    def _restyle_active_tooltip(self) -> None:
        app = QtWidgets.QApplication.instance()
        if app is None:
            return
        tooltip = app.findChild(QtWidgets.QLabel, "qt_tip_label")
        if tooltip is None:
            tooltip = app.findChild(QtWidgets.QLabel, "qt_tooltip_label")
        if tooltip is not None:
            self._apply_styles(tooltip)

    def _apply_styles(self, tooltip: QtWidgets.QWidget) -> None:
        tooltip.setAttribute(QtCore.Qt.WA_StyledBackground, True)
        left, right = self._padding
        tooltip.setStyleSheet(
            "QLabel {"
            f" color: {self._text};"
            f" background-color: {self._background};"
            f" border: 1px solid {self._border.name(QtGui.QColor.HexArgb)};"
            " border-radius: 8px;"
            f" padding: 4px {right}px 4px {left}px;"
            "}"
        )
        shadow = tooltip.graphicsEffect()
        if not isinstance(shadow, QtWidgets.QGraphicsDropShadowEffect):
            shadow = QtWidgets.QGraphicsDropShadowEffect(tooltip)
            tooltip.setGraphicsEffect(shadow)
        shadow.setBlurRadius(18)
        shadow.setOffset(0, 4)
        shadow.setColor(self._shadow_color)


class ThemedTooltip(QtWidgets.QFrame):
    """Custom tooltip widget with rounded corners and drop shadow."""

    def __init__(self) -> None:
        super().__init__(parent=None)
        self.setWindowFlag(QtCore.Qt.ToolTip, True)
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint, True)
        self.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint, True)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self.setAttribute(QtCore.Qt.WA_ShowWithoutActivating, True)
        self._background = QtGui.QColor("#ffffff")
        self._border = QtGui.QColor(ACCENT_COLOR)
        self._text = QtGui.QColor("#1f1f23")

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._label = QtWidgets.QLabel()
        self._label.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self._label.setTextFormat(QtCore.Qt.PlainText)
        self._label.setWordWrap(True)
        layout.addWidget(self._label)

        shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(14)
        shadow.setOffset(0, 4)
        shadow.setColor(QtGui.QColor(31, 31, 35, 100))
        self.setGraphicsEffect(shadow)
        self._shadow_effect = shadow

    def set_palette(self, *, background: str, text: str, border: QtGui.QColor, shadow: QtGui.QColor) -> None:
        self._background = QtGui.QColor(background)
        self._text = QtGui.QColor(text)
        self._border = border
        self._shadow_effect.setColor(shadow)
        self._label.setStyleSheet(f"QLabel {{ color: {self._text.name()}; padding: 6px 10px; }}")
        self.update()

    def set_text(self, text: str) -> None:
        self._label.setText(text)
        self.adjustSize()

    def show_at(self, pos: QtCore.QPoint) -> None:
        screen = QtGui.QGuiApplication.screenAt(pos)
        geom = screen.availableGeometry() if screen else QtGui.QGuiApplication.primaryScreen().availableGeometry()
        self.adjustSize()
        size = self.size()
        x = pos.x() + 12
        y = pos.y() + 16
        if x + size.width() > geom.right():
            x = geom.right() - size.width() - 8
        if y + size.height() > geom.bottom():
            y = pos.y() - size.height() - 12
        self.move(x, y)
        self.show()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:  # noqa: N802
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        rect = self.rect().adjusted(4, 4, -4, -4)
        path = QtGui.QPainterPath()
        radius = 2.0
        path.addRoundedRect(rect, radius, radius)
        painter.setPen(QtGui.QPen(self._border, 1))
        painter.setBrush(QtGui.QBrush(self._background))
        painter.drawPath(path)


class CustomTooltipManager(QtCore.QObject):
    """Global manager that displays themed tooltips and suppresses system styling."""

    def __init__(self, parent: QtCore.QObject | None = None) -> None:
        super().__init__(parent)
        self._tooltip_widget = ThemedTooltip()
        self._tooltip_widget.hide()
        app = QtWidgets.QApplication.instance()
        if app is not None:
            app.installEventFilter(self)
        self._hide_timer = QtCore.QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.timeout.connect(self._tooltip_widget.hide)

    def update_theme(self, *, background: str, text: str, border: QtGui.QColor, shadow: QtGui.QColor) -> None:
        self._tooltip_widget.set_palette(background=background, text=text, border=border, shadow=shadow)

    def show_text(self, text: str, global_pos: QtCore.QPoint) -> None:
        if not text:
            self._tooltip_widget.hide()
            return
        self._tooltip_widget.set_text(text)
        self._tooltip_widget.show_at(global_pos)
        self._hide_timer.start(8000)

    def hide(self) -> None:
        self._hide_timer.stop()
        self._tooltip_widget.hide()

    def eventFilter(self, obj: QtCore.QObject, event: QtCore.QEvent) -> bool:  # noqa: N802
        if event.type() == QtCore.QEvent.ToolTip:
            widget = obj if isinstance(obj, QtWidgets.QWidget) else None
            help_event = event  # type: ignore[assignment]
            text = widget.toolTip() if widget is not None else ""
            if text:
                self.show_text(text, help_event.globalPos())
            else:
                self.hide()
            return True
        if event.type() in {
            QtCore.QEvent.Leave,
            QtCore.QEvent.FocusOut,
            QtCore.QEvent.WindowDeactivate,
            QtCore.QEvent.HoverLeave,
            QtCore.QEvent.MouseButtonPress,
        }:
            self.hide()
        return super().eventFilter(obj, event)

class SpreadsheetPanel(QtWidgets.QWidget):
    """Left panel showing preview of the tabular dataset."""

    def __init__(self) -> None:
        super().__init__()
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(6)

        self.columns_widget = ColumnListWidget()
        self.columns_widget.setMinimumHeight(120)
        self.setMinimumWidth(260)
        self.setMaximumWidth(520)
        size_policy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        self.setSizePolicy(size_policy)

        summary_layout = QtWidgets.QHBoxLayout()
        summary_layout.addStretch()
        self.data_summary_label = QtWidgets.QLabel("No dataset loaded")
        self.data_summary_label.setObjectName("dataSummaryLabel")
        self.data_summary_label.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        summary_layout.addWidget(self.data_summary_label)

        column_container = QtWidgets.QWidget()
        column_layout = QtWidgets.QVBoxLayout(column_container)
        column_layout.setContentsMargins(0, 0, 0, 0)
        column_layout.setSpacing(4)
        column_header = QtWidgets.QLabel("Columns")
        column_layout.addWidget(column_header)
        self.column_search = QtWidgets.QLineEdit()
        self.column_search.setPlaceholderText("Search columns...")
        self.column_search.setClearButtonEnabled(True)
        self.column_search.setEnabled(False)
        self.column_search.hide()
        column_layout.addWidget(self.column_search)
        column_layout.addWidget(self.columns_widget)
        self.table_model = DataFrameModel()
        self.table_view = QtWidgets.QTableView()
        self.table_view.setModel(self.table_model)
        self.table_view.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.table_view.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table_view.horizontalHeader().setStretchLastSection(True)
        self.table_view.verticalHeader().setVisible(False)

        table_container = QtWidgets.QWidget()
        table_layout = QtWidgets.QVBoxLayout(table_container)
        table_layout.setContentsMargins(0, 0, 0, 0)
        table_layout.setSpacing(4)
        table_layout.addWidget(QtWidgets.QLabel("Sample Rows"))
        table_layout.addWidget(self.table_view)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        splitter.addWidget(column_container)
        splitter.addWidget(table_container)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

        self.column_search.textChanged.connect(self.columns_widget.apply_filter)

        layout.addLayout(summary_layout)
        layout.addWidget(splitter)

    def set_data(self, sample: DataSample) -> None:
        columns = sample.columns()
        if self.column_search.text():
            self.column_search.setText("")
        has_columns = bool(columns)
        self.column_search.setEnabled(has_columns)
        self.column_search.setVisible(has_columns)
        self.columns_widget.set_columns(columns)
        self.columns_widget.clearSelection()
        preview = sample.head_records(50)
        self.table_model.update(preview)
        rows, cols = sample.dataframe.shape
        sheet_suffix = f" | Sheet: {sample.sheet_name}" if sample.sheet_name else ""
        self.data_summary_label.setText(f"{rows:,} rows x {cols} columns{sheet_suffix}")

    def clear(self) -> None:
        self.columns_widget.clear()
        self.columns_widget.clearSelection()
        self.column_search.clear()
        self.column_search.setEnabled(False)
        self.column_search.hide()
        self.table_model.update(pd.DataFrame())
        self.data_summary_label.setText("No dataset loaded")


class DataRangeDialog(QtWidgets.QDialog):
    """Dialog allowing users to adjust header/data rows and column offset."""

    def __init__(
        self,
        parent: Optional[QtWidgets.QWidget] = None,
        *,
        header_row: int = 1,
        data_row: int = 2,
        first_column: int = 1,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Adjust Data Range")
        self.setModal(True)

        layout = QtWidgets.QVBoxLayout(self)
        form = QtWidgets.QFormLayout()

        self.header_spin = QtWidgets.QSpinBox()
        self.header_spin.setMinimum(1)
        self.header_spin.setMaximum(100000)
        self.header_spin.setValue(header_row)

        self.data_spin = QtWidgets.QSpinBox()
        self.data_spin.setMinimum(max(1, data_row))
        self.data_spin.setMaximum(100000)
        self.data_spin.setValue(max(data_row, header_row + 1))

        self.column_spin = QtWidgets.QSpinBox()
        self.column_spin.setMinimum(1)
        self.column_spin.setMaximum(1000)
        self.column_spin.setValue(max(1, first_column))

        form.addRow("Header row (1-indexed):", self.header_spin)
        form.addRow("First data row:", self.data_spin)
        form.addRow("First data column:", self.column_spin)
        layout.addLayout(form)

        hint = QtWidgets.QLabel("Use this dialog when column titles or data start lower in the file.")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.header_spin.valueChanged.connect(self._sync_data_minimum)
        self._sync_data_minimum(self.header_spin.value())

    def _sync_data_minimum(self, header_value: int) -> None:
        self.data_spin.setMinimum(header_value + 1)

    def values(self) -> tuple[int, int, int]:
        return (
            self.header_spin.value(),
            self.data_spin.value(),
            self.column_spin.value(),
        )


class GeneratePdfDialog(QtWidgets.QDialog):
    """Prompt the user for PDF generation options."""

    def __init__(
        self,
        parent: QtWidgets.QWidget | None,
        *,
        columns: Sequence[str],
        sample_row: Mapping[str, Any] | None,
        default_mode: str = "per_entry",
        default_directory: Path | None = None,
        default_file: Path | None = None,
        default_columns: Sequence[str] | None = None,
        default_prefix: str = "",
        default_suffix: str = "",
        default_separator: str = "_",
        default_read_only: bool = False,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Generate PDFs")
        self.setModal(True)
        self.resize(520, 520)

        self._all_columns = list(columns)
        self._sample_row = sample_row or {}
        self._default_mode = default_mode if default_mode in {"per_entry", "combined"} else "per_entry"

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(12)

        mode_group = QtWidgets.QGroupBox("Output Type")
        mode_layout = QtWidgets.QHBoxLayout(mode_group)
        self._per_entry_radio = QtWidgets.QRadioButton("One PDF per data row")
        self._combined_radio = QtWidgets.QRadioButton("Single combined PDF")
        mode_layout.addWidget(self._per_entry_radio)
        mode_layout.addWidget(self._combined_radio)
        mode_layout.addStretch(1)
        layout.addWidget(mode_group)

        self._stack = QtWidgets.QStackedWidget()
        layout.addWidget(self._stack)

        # --- Per-entry widget
        per_entry_widget = QtWidgets.QWidget()
        per_layout = QtWidgets.QVBoxLayout(per_entry_widget)
        per_layout.setSpacing(10)

        dir_layout = QtWidgets.QHBoxLayout()
        dir_layout.addWidget(QtWidgets.QLabel("Destination folder:"))
        self._directory_edit = QtWidgets.QLineEdit()
        self._directory_edit.setReadOnly(True)
        if default_directory:
            self._directory_edit.setText(str(default_directory))
        dir_layout.addWidget(self._directory_edit, 1)
        browse_dir_button = QtWidgets.QToolButton()
        browse_dir_button.setText("Browse…")
        browse_dir_button.clicked.connect(self._choose_directory)
        dir_layout.addWidget(browse_dir_button)
        per_layout.addLayout(dir_layout)

        column_group = QtWidgets.QGroupBox("Filename columns")
        column_group.setToolTip("Select columns whose values should appear in each PDF filename.")
        column_layout = QtWidgets.QVBoxLayout(column_group)
        self._column_list = QtWidgets.QListWidget()
        self._column_list.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        column_layout.addWidget(self._column_list)
        per_layout.addWidget(column_group)

        options_layout = QtWidgets.QGridLayout()
        options_layout.addWidget(QtWidgets.QLabel("Prefix:"), 0, 0)
        self._prefix_edit = QtWidgets.QLineEdit(default_prefix)
        options_layout.addWidget(self._prefix_edit, 0, 1)

        options_layout.addWidget(QtWidgets.QLabel("Suffix:"), 1, 0)
        self._suffix_edit = QtWidgets.QLineEdit(default_suffix)
        options_layout.addWidget(self._suffix_edit, 1, 1)

        options_layout.addWidget(QtWidgets.QLabel("Separator:"), 2, 0)
        self._separator_edit = QtWidgets.QLineEdit(default_separator or "_")
        self._separator_edit.setMaxLength(4)
        options_layout.addWidget(self._separator_edit, 2, 1)
        per_layout.addLayout(options_layout)

        self._filename_preview = QtWidgets.QLabel("")
        self._filename_preview.setObjectName("filenamePreviewLabel")
        per_layout.addWidget(self._filename_preview)

        per_layout.addStretch(1)
        self._stack.addWidget(per_entry_widget)

        # --- Combined widget
        combined_widget = QtWidgets.QWidget()
        combined_layout = QtWidgets.QFormLayout(combined_widget)
        self._combined_path_edit = QtWidgets.QLineEdit()
        if default_file:
            self._combined_path_edit.setText(str(default_file))
        combined_browse = QtWidgets.QToolButton()
        combined_browse.setText("Browse…")
        combined_browse.clicked.connect(self._choose_combined_path)

        file_layout = QtWidgets.QHBoxLayout()
        file_layout.addWidget(self._combined_path_edit, 1)
        file_layout.addWidget(combined_browse)
        combined_layout.addRow("Output file:", file_layout)
        note = QtWidgets.QLabel("The generated PDFs will be merged into a single document.")
        note.setWordWrap(True)
        combined_layout.addRow(note)
        combined_layout.addItem(QtWidgets.QSpacerItem(10, 10, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding))
        self._stack.addWidget(combined_widget)

        # --- Output behavior
        behavior_group = QtWidgets.QGroupBox("PDF behavior")
        behavior_layout = QtWidgets.QHBoxLayout(behavior_group)
        self._editable_radio = QtWidgets.QRadioButton("Editable (fillable)")
        self._readonly_radio = QtWidgets.QRadioButton("Read-only (locked)")
        behavior_layout.addWidget(self._editable_radio)
        behavior_layout.addWidget(self._readonly_radio)
        behavior_layout.addStretch(1)
        layout.addWidget(behavior_group)

        # Buttons
        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        layout.addWidget(buttons)

        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        # Populate columns list
        for column in self._all_columns:
            item = QtWidgets.QListWidgetItem(column)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            item.setCheckState(QtCore.Qt.Unchecked)
            self._column_list.addItem(item)

        if default_columns:
            default_set = {col for col in default_columns}
            for index in range(self._column_list.count()):
                item = self._column_list.item(index)
                if item.text() in default_set:
                    item.setCheckState(QtCore.Qt.Checked)

        self._per_entry_radio.toggled.connect(self._update_mode)
        self._column_list.itemChanged.connect(self._update_preview)
        self._prefix_edit.textChanged.connect(self._update_preview)
        self._suffix_edit.textChanged.connect(self._update_preview)
        self._separator_edit.textChanged.connect(self._update_preview)

        if default_read_only:
            self._readonly_radio.setChecked(True)
        else:
            self._editable_radio.setChecked(True)

        if self._default_mode == "combined":
            self._combined_radio.setChecked(True)
        else:
            self._per_entry_radio.setChecked(True)

        self._update_mode()
        self._update_preview()

    def selected_columns(self) -> list[str]:
        return [
            self._column_list.item(i).text()
            for i in range(self._column_list.count())
            if self._column_list.item(i).checkState() == QtCore.Qt.Checked
        ]

    def options(self) -> GenerationOptions:
        mode = "combined" if self._combined_radio.isChecked() else "per_entry"
        directory = Path(self._directory_edit.text()) if self._directory_edit.text() else None
        combined_path = (
            Path(self._combined_path_edit.text()) if self._combined_path_edit.text() else None
        )
        separator = self._separator_edit.text() or "_"
        return GenerationOptions(
            mode=mode,
            destination_dir=directory,
            combined_path=combined_path,
            columns=self.selected_columns(),
            prefix=self._prefix_edit.text(),
            suffix=self._suffix_edit.text(),
            separator=separator,
            read_only=self._readonly_radio.isChecked(),
        )

    def accept(self) -> None:  # noqa: D401
        """Validate user choices before closing the dialog."""
        mode = "combined" if self._combined_radio.isChecked() else "per_entry"
        if mode == "per_entry":
            directory = self._directory_edit.text()
            if not directory:
                QtWidgets.QMessageBox.warning(self, "Missing folder", "Select an output folder.")
                return
        else:
            path_text = self._combined_path_edit.text()
            if not path_text:
                QtWidgets.QMessageBox.warning(self, "Missing file", "Choose a combined PDF filename.")
                return
            if not path_text.lower().endswith(".pdf"):
                self._combined_path_edit.setText(f"{path_text}.pdf")
        super().accept()

    def _choose_directory(self) -> None:
        current = self._directory_edit.text()
        path = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select output folder", current or ""
        )
        if path:
            self._directory_edit.setText(path)
            self._update_preview()

    def _choose_combined_path(self) -> None:
        current = self._combined_path_edit.text()
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save merged PDF",
            current or "",
            "PDF Files (*.pdf)",
        )
        if path:
            if not path.lower().endswith(".pdf"):
                path = f"{path}.pdf"
            self._combined_path_edit.setText(path)

    def _update_mode(self) -> None:
        per_entry = self._per_entry_radio.isChecked()
        self._stack.setCurrentIndex(0 if per_entry else 1)
        self._stack.setEnabled(True)
        self._column_list.setEnabled(per_entry)
        self._prefix_edit.setEnabled(per_entry)
        self._suffix_edit.setEnabled(per_entry)
        self._separator_edit.setEnabled(per_entry)
        self._update_preview()

    def _update_preview(self) -> None:
        if not self._per_entry_radio.isChecked():
            self._filename_preview.setText("Files will be merged into a single PDF.")
            return
        if not self._sample_row:
            self._filename_preview.setText("Preview unavailable (data sample not loaded).")
            return
        builder = FilenameBuilder(
            self.selected_columns(),
            prefix=self._prefix_edit.text(),
            suffix=self._suffix_edit.text(),
            separator=self._separator_edit.text() or "_",
        )
        preview = builder.preview(self._sample_row, 1)
        self._filename_preview.setText(f"Example filename: {preview}.pdf")


class PdfFieldItem(QtWidgets.QGraphicsRectItem):
    """Interactive overlay representing a PDF form field on the canvas."""

    def __init__(
        self,
        field: PdfField,
        rect: QtCore.QRectF,
        drop_callback: Callable[[PdfField, str], None],
        click_callback: Callable[[PdfField], None],
        remove_callback: Callable[[PdfField], None],
        tooltip_manager: CustomTooltipManager | None,
    ) -> None:
        super().__init__(rect)
        self.field = field
        self._drop_callback = drop_callback
        self._click_callback = click_callback
        self._remove_callback = remove_callback
        self._tooltip_manager = tooltip_manager
        self.setBrush(QtGui.QColor(0, 170, 255, 50))
        self._base_pen = QtGui.QPen(QtGui.QColor(0, 120, 215), 1, QtCore.Qt.DashLine)
        self._selected_pen = QtGui.QPen(QtGui.QColor(ACCENT_COLOR), 2)
        self._selected_pen.setStyle(QtCore.Qt.SolidLine)
        self.setPen(self._base_pen)
        self.setZValue(1)
        self.setAcceptDrops(True)
        self.setAcceptHoverEvents(True)
        self.setAcceptedMouseButtons(QtCore.Qt.AllButtons)
        self._label: QtWidgets.QGraphicsSimpleTextItem | None = None
        self._current_column: str | None = None
        self._selected = False
        self._overlay_container = QtWidgets.QWidget()
        overlay_layout = QtWidgets.QHBoxLayout(self._overlay_container)
        overlay_layout.setContentsMargins(2, 2, 2, 2)
        overlay_layout.setSpacing(4)

        self._select_button = QtWidgets.QToolButton()
        self._select_button.setIcon(get_fluent_icon("SELECT_ALL", "ADD", default=FI.ADD))
        self._select_button.setIconSize(QtCore.QSize(16, 16))
        self._select_button.setToolTip("Select this field")
        self._select_button.setAutoRaise(True)
        self._select_button.clicked.connect(lambda: self._click_callback(self.field))

        self._edit_button = QtWidgets.QToolButton()
        self._edit_button.setIcon(get_fluent_icon("EDIT", default=FI.EDIT))
        self._edit_button.setIconSize(QtCore.QSize(16, 16))
        self._edit_button.setToolTip("Edit mapping")
        self._edit_button.setAutoRaise(True)
        self._edit_button.clicked.connect(lambda: self._click_callback(self.field))

        self._remove_button = QtWidgets.QToolButton()
        self._remove_button.setIcon(get_fluent_icon("DELETE", default=FI.DELETE))
        self._remove_button.setIconSize(QtCore.QSize(16, 16))
        self._remove_button.setToolTip("Remove mapping")
        self._remove_button.setAutoRaise(True)
        self._remove_button.clicked.connect(lambda: self._remove_callback(self.field))

        overlay_layout.addWidget(self._select_button)
        overlay_layout.addWidget(self._edit_button)
        overlay_layout.addWidget(self._remove_button)

        self._overlay_proxy = QtWidgets.QGraphicsProxyWidget(self)
        self._overlay_proxy.setWidget(self._overlay_container)
        self._overlay_proxy.setZValue(3)
        self._overlay_proxy.setVisible(False)

    def dragEnterEvent(self, event: QtWidgets.QGraphicsSceneDragDropEvent) -> None:  # noqa: N802
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QtWidgets.QGraphicsSceneDragDropEvent) -> None:  # noqa: N802
        if event.mimeData().hasText():
            self._drop_callback(self.field, event.mimeData().text())
            event.acceptProposedAction()
        else:
            event.ignore()

    def update_assignment(
        self,
        column_name: Optional[str],
        sample_value: Optional[str] = None,
    ) -> None:
        self._current_column = column_name
        if column_name:
            self.setToolTip(f"PDF Field: {self.field.field_name}\nColumn: {column_name}")
            self.setBrush(QtGui.QColor(0, 170, 255, 100))
            self._set_label_text(sample_value or "")
        else:
            self.setToolTip(f"PDF Field: {self.field.field_name}")
            self.setBrush(QtGui.QColor(0, 170, 255, 50))
            if self._label:
                self._label.setText("")
        self._apply_selection_style()
        self._update_overlay()

    def _ensure_label(self) -> QtWidgets.QGraphicsSimpleTextItem:
        if self._label is None:
            self._label = QtWidgets.QGraphicsSimpleTextItem("", self)
            self._label.setBrush(QtGui.QBrush(QtGui.QColor(0, 0, 0)))
            self._label.setZValue(2)
            self._label.setPos(self.rect().left() + 2, self.rect().top() + 2)
            font = QtGui.QFont()
            font.setPointSize(8)
            self._label.setFont(font)
        return self._label
    def hoverEnterEvent(self, event: QtWidgets.QGraphicsSceneHoverEvent) -> None:  # noqa: N802
        self._show_tooltip(event)
        super().hoverEnterEvent(event)

    def hoverMoveEvent(self, event: QtWidgets.QGraphicsSceneHoverEvent) -> None:  # noqa: N802
        self._show_tooltip(event)
        super().hoverMoveEvent(event)

    def hoverLeaveEvent(self, event: QtWidgets.QGraphicsSceneHoverEvent) -> None:  # noqa: N802
        if self._tooltip_manager:
            self._tooltip_manager.hide()
        else:
            QtWidgets.QToolTip.hideText()
        super().hoverLeaveEvent(event)

    def mousePressEvent(self, event: QtWidgets.QGraphicsSceneMouseEvent) -> None:  # noqa: N802
        if event.button() == QtCore.Qt.LeftButton:
            self._click_callback(self.field)
        self._show_tooltip(event)
        super().mousePressEvent(event)

    def contextMenuEvent(self, event: QtWidgets.QGraphicsSceneContextMenuEvent) -> None:  # noqa: N802
        if self._tooltip_manager:
            self._tooltip_manager.hide()
        menu = QtWidgets.QMenu()
        if self._current_column:
            edit_icon = get_fluent_icon("EDIT", default=FI.EDIT)
            edit_action = menu.addAction(edit_icon, "Edit Mapping…")
            edit_action.triggered.connect(lambda: self._click_callback(self.field))
            remove_icon = get_fluent_icon("DELETE", default=FI.DELETE)
            remove_action = menu.addAction(remove_icon, "Remove Mapping")
            remove_action.triggered.connect(lambda: self._remove_callback(self.field))
        else:
            select_icon = get_fluent_icon("SELECT_ALL", "ADD", default=FI.ADD)
            select_action = menu.addAction(select_icon, "Select Field")
            select_action.triggered.connect(lambda: self._click_callback(self.field))
        menu.exec(event.screenPos())

    def _show_tooltip(self, event: QtCore.QEvent) -> None:
        tooltip = self.toolTip() or f"PDF Field: {self.field.field_name}"
        if isinstance(event, (QtWidgets.QGraphicsSceneHoverEvent, QtWidgets.QGraphicsSceneMouseEvent)):
            pos = event.screenPos()
        else:
            pos = QtGui.QCursor.pos()
        global_pos = QtCore.QPoint(int(pos.x()), int(pos.y()))
        if self._tooltip_manager:
            self._tooltip_manager.show_text(tooltip, global_pos)
        else:
            QtWidgets.QToolTip.showText(
                global_pos,
                tooltip,
                None,
            )

    def _set_label_text(self, text: str) -> None:
        label = self._ensure_label()
        if not text:
            label.setText("")
            return

        rect = self.rect()
        stripped = text.strip()
        is_checkbox_preview = stripped in {"\u2713", "✓", "✔", "☑"}

        if is_checkbox_preview:
            base = min(rect.width(), rect.height())
            preferred = max(8.0, min(base * 0.68, 14.0))
            font = label.font()
            font.setPointSizeF(preferred)
            label.setFont(font)

            metrics = QtGui.QFontMetricsF(font)
            display = stripped
            text_width = metrics.horizontalAdvance(display)
            text_height = metrics.height()
            x = rect.left() + max(0.0, (rect.width() - text_width) / 2.0)
            y = rect.top() + max(0.0, (rect.height() - text_height) / 2.0)
            label.setPos(x, y)
            label.setText(display)
            return

        horizontal_padding = max(2.0, min(rect.width() * 0.04, 6.0))
        vertical_padding = max(2.0, min(rect.height() * 0.2, 6.0))
        usable_width = max(8.0, rect.width() - (horizontal_padding * 2.0))
        usable_height = max(7.0, rect.height() - (vertical_padding * 2.0))

        font = label.font()
        base_size = min(11.0, max(8.0, usable_height * 0.6))
        font_size = self._calculate_font_size(
            font,
            text,
            usable_width,
            usable_height,
            base_size,
        )
        font.setPointSizeF(font_size)
        label.setFont(font)

        metrics = QtGui.QFontMetricsF(font)
        elided = metrics.elidedText(text, QtCore.Qt.ElideRight, usable_width)
        label.setPos(rect.left() + horizontal_padding, rect.top() + vertical_padding)
        label.setText(elided)

    def set_selected(self, selected: bool) -> None:
        if self._selected == selected:
            return
        self._selected = selected
        self._apply_selection_style()
        self._update_overlay()

    def _apply_selection_style(self) -> None:
        self.setPen(self._selected_pen if self._selected else self._base_pen)
        self._update_overlay()

    def _update_overlay(self) -> None:
        if self._overlay_proxy is None:
            return
        has_mapping = bool(self._current_column)
        if not self._selected:
            self._overlay_proxy.setVisible(False)
            return

        self._select_button.setVisible(not has_mapping)
        self._edit_button.setVisible(has_mapping)
        self._remove_button.setVisible(has_mapping)
        self._overlay_container.adjustSize()
        self._position_overlay()
        self._overlay_proxy.setVisible(True)

    def _position_overlay(self) -> None:
        if not self._overlay_proxy or not self._overlay_proxy.widget():
            return
        rect = self.rect()
        size = self._overlay_container.sizeHint()
        x = rect.right() - size.width() - 4.0
        y = rect.top() + 4.0
        self._overlay_proxy.setPos(x, y)

    def _calculate_font_size(
        self,
        font: QtGui.QFont,
        text: str,
        max_width: float,
        max_height: float,
        base_size: float,
    ) -> float:
        """Return a font size that fits within the provided bounds."""
        min_size = 6.0
        size = max(base_size, min_size)
        font.setPointSizeF(size)
        metrics = QtGui.QFontMetricsF(font)
        while size > min_size and (
            metrics.height() > max_height or metrics.horizontalAdvance(text) > max_width
        ):
            size -= 0.5
            font.setPointSizeF(size)
            metrics = QtGui.QFontMetricsF(font)
        return max(size, min_size)


class PdfViewerWidget(QtWidgets.QGraphicsView):
    """Displays a rendered PDF page with draggable form fields."""

    fieldAssigned = QtCore.Signal(str, str)
    fieldActivated = QtCore.Signal(str)
    fieldRemoveRequested = QtCore.Signal(str)
    fieldSelectionChanged = QtCore.Signal(object)
    pageChanged = QtCore.Signal(int)
    zoomChanged = QtCore.Signal(float)

    def __init__(self, engine: PdfEngine, *, tooltip_manager: CustomTooltipManager | None = None) -> None:
        super().__init__()
        self._engine = engine
        self._template: PdfTemplate | None = None
        self._zoom = 1.0
        self._current_page = 0
        self._page_count = 0
        self._field_items: Dict[str, PdfFieldItem] = {}
        self._auto_fit = False
        self._assignments: Dict[str, tuple[str | None, str | None]] = {}
        self._selected_field: Optional[str] = None
        self._tooltip_manager = tooltip_manager

        self.setScene(QtWidgets.QGraphicsScene(self))
        self.setRenderHint(QtGui.QPainter.Antialiasing)
        self.setAcceptDrops(True)
        self.setAlignment(QtCore.Qt.AlignCenter)

    def load_template(self, template: PdfTemplate, page_index: int = 0) -> None:
        self._template = template
        self._assignments.clear()
        self._page_count = template.document.page_count
        self._current_page = max(0, min(page_index, self._page_count - 1))
        self._set_selected_field(None)
        self._render_page()
        self.pageChanged.emit(self._current_page)
        self.zoomChanged.emit(self._zoom)

    def clear(self) -> None:
        self.scene().clear()
        self._field_items.clear()
        self._template = None
        self._page_count = 0
        self._current_page = 0
        self._auto_fit = False
        self._assignments.clear()
        self._set_selected_field(None)
        self.zoomChanged.emit(self._zoom)

    def _handle_drop(self, field: PdfField, column_name: str) -> None:
        self.fieldAssigned.emit(field.field_name, column_name)
        self._set_selected_field(None)

    def set_assignment(
        self,
        field_name: str,
        column_name: Optional[str],
        sample_value: Optional[str] = None,
    ) -> None:
        if item := self._field_items.get(field_name):
            item.update_assignment(column_name, sample_value)
            item.set_selected(field_name == self._selected_field)
        if column_name:
            self._assignments[field_name] = (column_name, sample_value)
        else:
            self._assignments.pop(field_name, None)

    def clear_assignments(self) -> None:
        """Remove all visual assignment markers from the viewer."""
        for field_name in list(self._assignments.keys()):
            self.set_assignment(field_name, None, None)
        self._assignments.clear()
        self._set_selected_field(None)

    def set_page(self, page_index: int) -> None:
        if self._template is None:
            return
        normalized = max(0, min(page_index, self._page_count - 1))
        if normalized == self._current_page:
            return
        self._current_page = normalized
        self._render_page()
        self.pageChanged.emit(self._current_page)
        self.zoomChanged.emit(self._zoom)

    def current_page(self) -> int:
        return self._current_page

    def page_count(self) -> int:
        return self._page_count

    def _render_page(self) -> None:
        if self._template is None:
            self.scene().clear()
            self._field_items.clear()
            return

        self.scene().clear()
        self._field_items.clear()

        pixmap = self._engine.render_page(self._template, self._current_page, zoom=self._zoom)
        image = QtGui.QImage(
            pixmap.samples, pixmap.width, pixmap.height, pixmap.stride, QtGui.QImage.Format_RGB888
        )
        qt_pixmap = QtGui.QPixmap.fromImage(image)
        self.scene().addPixmap(qt_pixmap)

        transform = fitz.Matrix(self._zoom, self._zoom)
        for field in self._template.fields:
            if field.page_index != self._current_page:
                continue
            scaled = field.rect * transform
            rect = QtCore.QRectF(scaled.x0, scaled.y0, scaled.width, scaled.height)
            item = PdfFieldItem(
                field,
                rect,
                self._handle_drop,
                self._handle_field_click,
                self._handle_field_remove,
                self._tooltip_manager,
            )
            self.scene().addItem(item)
            self._field_items[field.field_name] = item
            item.set_selected(field.field_name == self._selected_field)

        for field_name, (column, preview) in self._assignments.items():
            if item := self._field_items.get(field_name):
                item.update_assignment(column, preview)
                item.set_selected(field_name == self._selected_field)

        self.scene().setSceneRect(self.scene().itemsBoundingRect())
        self.resetTransform()
        self.centerOn(self.scene().sceneRect().center())

    def wheelEvent(self, event: QtGui.QWheelEvent) -> None:  # noqa: N802
        if event.modifiers() & QtCore.Qt.ControlModifier:
            angle_delta = event.angleDelta()
            delta = angle_delta.y() or angle_delta.x()
            if delta != 0:
                steps = delta / 120.0
                factor = 1.25 ** steps
                self.set_zoom(self._zoom * factor, auto_fit=False)
            event.accept()
            return
        super().wheelEvent(event)

    def set_zoom(self, zoom: float, *, auto_fit: bool = False) -> None:
        zoom = max(0.25, min(5.0, zoom))
        self._zoom = zoom
        self._auto_fit = auto_fit
        if self._template:
            self._render_page()
        self.zoomChanged.emit(self._zoom)

    def zoom_in(self) -> None:
        self.set_zoom(self._zoom * 1.25)

    def zoom_out(self) -> None:
        self.set_zoom(self._zoom / 1.25)

    def fit_to_width(self) -> None:
        if not self._template:
            return
        viewport_width = self.viewport().width()
        if viewport_width <= 0:
            return
        page = self._template.document.load_page(self._current_page)
        page_width = page.rect.width
        if page_width > 0:
            fit_zoom = viewport_width / page_width
            self.set_zoom(fit_zoom, auto_fit=True)

    def actual_size(self) -> None:
        self.set_zoom(1.0, auto_fit=False)

    def current_zoom(self) -> float:
        return self._zoom

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:  # noqa: N802
        super().resizeEvent(event)
        if self._auto_fit:
            self.fit_to_width()
        if self._template is None:
            return

        self.centerOn(self.scene().sceneRect().center())

    def clear_field_selection(self) -> None:
        self._set_selected_field(None)

    def _handle_field_remove(self, field: PdfField) -> None:
        self.fieldRemoveRequested.emit(field.field_name)
        self._set_selected_field(None)

    def _handle_field_click(self, field: PdfField) -> None:
        field_name = field.field_name
        if field_name in self._assignments and self._assignments[field_name][0]:
            self.fieldActivated.emit(field_name)
            return
        if self._selected_field == field_name:
            self._set_selected_field(None)
        else:
            self._set_selected_field(field_name)

    def _set_selected_field(self, field_name: Optional[str], *, emit: bool = True) -> None:
        if field_name == self._selected_field:
            return
        self._selected_field = field_name
        for name, item in self._field_items.items():
            item.set_selected(name == self._selected_field)
        if emit:
            self.fieldSelectionChanged.emit(self._selected_field)


class MappingTable(QtWidgets.QTableWidget):
    """Tabular display of current mappings."""

    editRequested = QtCore.Signal(str)
    removeRequested = QtCore.Signal(str)

    def __init__(self) -> None:
        super().__init__(0, 5)
        self.setHorizontalHeaderLabels(["Rule", "Targets", "Summary", "Preview", "Actions"])
        header = self.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)
        self.verticalHeader().setVisible(False)
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.setAlternatingRowColors(True)

    def update_mapping(
        self,
        assignments: Dict[str, MappingRule],
        previews: Dict[str, str] | None = None,
    ) -> None:
        previews = previews or {}
        self.setRowCount(len(assignments))
        for row, (field, rule) in enumerate(sorted(assignments.items(), key=lambda item: item[0])):
            field_item = QtWidgets.QTableWidgetItem(field)
            field_item.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            targets_item = QtWidgets.QTableWidgetItem(", ".join(rule.targets))
            targets_item.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            summary_item = QtWidgets.QTableWidgetItem(rule.describe())
            summary_item.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            preview_text = previews.get(field, "")
            preview_item = QtWidgets.QTableWidgetItem(preview_text)
            preview_item.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)

            self.setItem(row, 0, field_item)
            self.setItem(row, 1, targets_item)
            self.setItem(row, 2, summary_item)
            self.setItem(row, 3, preview_item)

            action_widget = QtWidgets.QWidget()
            action_layout = QtWidgets.QHBoxLayout(action_widget)
            action_layout.setContentsMargins(0, 0, 0, 0)
            action_layout.setSpacing(4)

            edit_button = QtWidgets.QToolButton()
            edit_button.setIcon(get_fluent_icon("EDIT"))
            edit_button.setIconSize(QtCore.QSize(16, 16))
            edit_button.setToolTip(f"Edit rule for {field}")
            edit_button.clicked.connect(lambda checked=False, f=field: self.editRequested.emit(f))

            remove_button = QtWidgets.QToolButton()
            remove_button.setIcon(get_fluent_icon("DELETE"))
            remove_button.setIconSize(QtCore.QSize(16, 16))
            remove_button.setToolTip(f"Remove mapping for {field}")
            remove_button.clicked.connect(lambda checked=False, f=field: self.removeRequested.emit(f))

            action_layout.addWidget(edit_button)
            action_layout.addWidget(remove_button)
            action_layout.addStretch(1)
            self.setCellWidget(row, 4, action_widget)

    def selected_field(self) -> Optional[str]:
        row = self.currentRow()
        if row < 0:
            return None
        item = self.item(row, 0)
        return item.text() if item else None

    def keyPressEvent(self, event: QtGui.QKeyEvent) -> None:  # noqa: N802
        if event.key() in (QtCore.Qt.Key_Delete, QtCore.Qt.Key_Backspace):
            field = self.selected_field()
            if field:
                self.removeRequested.emit(field)
            event.accept()
            return
        if event.key() in (QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter):
            field = self.selected_field()
            if field:
                self.editRequested.emit(field)
            event.accept()
            return
        super().keyPressEvent(event)

    def mouseDoubleClickEvent(self, event: QtGui.QMouseEvent) -> None:  # noqa: N802
        super().mouseDoubleClickEvent(event)
        field = self.selected_field()
        if field:
            self.editRequested.emit(field)


@dataclass
class UiState:
    """Track the current working objects for the session."""

    data_sample: DataSample | None = None
    pdf_template: PdfTemplate | None = None
    mapping: MappingModel = field(default_factory=MappingModel)


class MainWindow(QtWidgets.QMainWindow):
    """Primary application window."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("PDF Bulk Filler")
        self.resize(1400, 900)
        self.setAutoFillBackground(True)

        self._settings = QtCore.QSettings()
        self._theme_mode = str(self._settings.value(THEME_SETTINGS_KEY, DEFAULT_THEME_MODE))
        if self._theme_mode not in THEME_MAP:
            self._theme_mode = DEFAULT_THEME_MODE
        self._tooltip_manager: CustomTooltipManager | None = None
        self._tooltip_styler = TooltipStyler(self)
        # Added theme switcher and persistence
        self._apply_theme(self._theme_mode, save=False, update_actions=False)

        self._data_loader = DataLoader()
        self._pdf_engine = PdfEngine()
        self._mapping_manager = MappingManager()

        self._state = UiState()

        self.spreadsheet_panel = SpreadsheetPanel()
        self.spreadsheet_panel.columns_widget.columnActivated.connect(self._on_column_activated)
        self.pdf_viewer = PdfViewerWidget(self._pdf_engine, tooltip_manager=self._tooltip_manager)
        self.pdf_viewer.fieldAssigned.connect(self._on_field_assigned)
        self.pdf_viewer.fieldActivated.connect(lambda field: self._action_edit_mapping(field))
        self.pdf_viewer.fieldRemoveRequested.connect(self._on_field_remove_requested)
        self.pdf_viewer.fieldSelectionChanged.connect(self._on_field_selection_changed)
        self.pdf_viewer.pageChanged.connect(self._on_page_changed)
        self.pdf_viewer.zoomChanged.connect(self._on_zoom_changed)
        self._selected_viewer_field: Optional[str] = None

        self._pdf_placeholder = QtWidgets.QLabel(
            "Import a PDF template to preview form fields and begin mapping."
        )
        self._pdf_placeholder.setAlignment(QtCore.Qt.AlignCenter)
        self._pdf_placeholder.setWordWrap(True)
        self._pdf_placeholder.setObjectName("pdfPlaceholder")

        self.viewer_stack = QtWidgets.QStackedWidget()
        self.viewer_stack.addWidget(self._pdf_placeholder)
        self.viewer_stack.addWidget(self.pdf_viewer)
        self.viewer_stack.setCurrentWidget(self._pdf_placeholder)
        self.viewer_stack.setMinimumWidth(720)
        self.viewer_stack.setSizePolicy(
            QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        )

        self.mapping_table = MappingTable()
        self.mapping_table.editRequested.connect(lambda field: self._action_edit_mapping(field))
        self.mapping_table.removeRequested.connect(lambda field: self._action_remove_mapping(field))
        self.mapping_table.itemSelectionChanged.connect(self._update_mapping_action_state)
        self._mapping_dock = QtWidgets.QDockWidget("Mappings", self)
        self._mapping_dock.setWidget(self.mapping_table)
        self._mapping_dock.setAllowedAreas(QtCore.Qt.BottomDockWidgetArea | QtCore.Qt.TopDockWidgetArea)

        splitter = QtWidgets.QSplitter()
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self.spreadsheet_panel)
        splitter.addWidget(self.viewer_stack)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 4)
        splitter.setSizes([360, 1080])
        self._splitter = splitter

        container = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(splitter)
        self.setCentralWidget(container)
        self.addDockWidget(QtCore.Qt.BottomDockWidgetArea, self._mapping_dock)
        self._mapping_dock.hide()
        self._mapping_dock.toggleViewAction().setChecked(False)

        self._register_actions()
        self._create_menus()
        self._generation_thread: QtCore.QThread | None = None
        self._generation_worker: PdfGenerationWorker | None = None
        self._generation_progress: QtWidgets.QProgressDialog | None = None
        self._configure_page_controls_for_template()
        self._zoom_label = QtWidgets.QLabel("100%")
        self.statusBar().addPermanentWidget(self._zoom_label)
        self._read_only_output = False
        self._output_mode_label = QtWidgets.QLabel("Output Mode: Editable (fillable)")
        self.statusBar().addPermanentWidget(self._output_mode_label)
        self._preview_data_label = QtWidgets.QLabel("Preview data: none loaded")
        self._preview_data_label.setToolTip("Indicates which dataset row feeds the live PDF preview.")
        self.statusBar().addPermanentWidget(self._preview_data_label)
        self._last_generation_mode = "per_entry"
        self._last_generation_target: Path | None = None
        self._last_generation_read_only = False
        self._on_zoom_changed(self.pdf_viewer.current_zoom())
        self._update_preview_data_indicator()
        self._set_status("Load data and a PDF template to begin")

    # ----- Action configuration -------------------------------------------------
    def _register_actions(self) -> None:
        toolbar = QtWidgets.QToolBar("Main")
        toolbar.setMovable(False)
        # Improved toolbar icon sizing for visual consistency
        toolbar.setIconSize(QtCore.QSize(20, 20))

        self._import_data_action = QtGui.QAction("Import CSV/Excel", self)
        self._import_data_action.triggered.connect(self._action_import_data)
        self._import_data_action.setShortcut(QtGui.QKeySequence("Ctrl+Shift+D"))
        # Improved icon contrast with Fluent icons
        self._import_data_action.setIcon(get_fluent_icon("FOLDER_ADD", "CLOUD_DOWNLOAD", "DOWNLOAD", default=FI.DOWNLOAD))

        self._import_pdf_action = QtGui.QAction("Import PDF", self)
        self._import_pdf_action.triggered.connect(self._action_import_pdf)
        self._import_pdf_action.setShortcut(QtGui.QKeySequence("Ctrl+Shift+P"))
        # Improved icon contrast with Fluent icons
        self._import_pdf_action.setIcon(get_fluent_icon("DOCUMENT", default=FI.FOLDER))

        self._save_mapping_action = QtGui.QAction("Save Mapping", self)
        self._save_mapping_action.triggered.connect(self._action_save_mapping)
        self._save_mapping_action.setShortcut(QtGui.QKeySequence.Save)
        # Improved icon contrast with Fluent icons
        self._save_mapping_action.setIcon(get_fluent_icon("SAVE"))

        self._load_mapping_action = QtGui.QAction("Load Mapping", self)
        self._load_mapping_action.triggered.connect(self._action_load_mapping)
        self._load_mapping_action.setShortcut(QtGui.QKeySequence.Open)
        # Improved icon contrast with Fluent icons
        self._load_mapping_action.setIcon(get_fluent_icon("OPEN_FOLDER", "FOLDER", default=FI.FOLDER))

        self._adjust_range_action = QtGui.QAction("Adjust Data Range", self)
        self._adjust_range_action.setShortcut(QtGui.QKeySequence("Ctrl+Shift+R"))
        # Improved icon contrast with Fluent icons
        self._adjust_range_action.setIcon(get_fluent_icon("SYNC"))
        self._adjust_range_action.triggered.connect(self._action_adjust_data_range)
        self._adjust_range_action.setEnabled(False)

        self._edit_mapping_action = QtGui.QAction("Edit Mapping", self)
        self._edit_mapping_action.triggered.connect(lambda: self._action_edit_mapping())
        self._edit_mapping_action.setEnabled(False)
        self._edit_mapping_action.setShortcut(QtGui.QKeySequence("Ctrl+E"))
        self._edit_mapping_action.setIcon(get_fluent_icon("EDIT"))

        self._remove_mapping_action = QtGui.QAction("Remove Mapping", self)
        self._remove_mapping_action.triggered.connect(lambda: self._action_remove_mapping())
        self._remove_mapping_action.setEnabled(False)
        self._remove_mapping_action.setShortcut(QtGui.QKeySequence.Delete)
        # Improved icon contrast with Fluent icons
        self._remove_mapping_action.setIcon(get_fluent_icon("DELETE"))

        self._generate_action = QtGui.QAction("Generate PDFs", self)
        self._generate_action.triggered.connect(self._action_generate_pdfs)
        self._generate_action.setShortcut(QtGui.QKeySequence("Ctrl+G"))
        # Improved icon contrast with Fluent icons
        self._generate_action.setIcon(get_fluent_icon("PLAY", "ARROW_RIGHT", default=FI.SYNC))

        toolbar.addActions(
            [
                self._import_data_action,
                self._import_pdf_action,
                self._save_mapping_action,
                self._load_mapping_action,
                self._adjust_range_action,
                self._edit_mapping_action,
                  self._remove_mapping_action,
                  self._generate_action,
              ]
          )

        spacer = QtWidgets.QWidget()
        spacer.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        toolbar.addWidget(spacer)
        toolbar.addSeparator()
        page_label = QtWidgets.QLabel("Page")
        page_label.setContentsMargins(8, 0, 4, 0)
        toolbar.addWidget(page_label)
        self.page_spinner = QtWidgets.QSpinBox()
        self.page_spinner.setMinimum(1)
        self.page_spinner.setMaximum(1)
        self.page_spinner.setEnabled(False)
        self.page_spinner.valueChanged.connect(self._on_page_selected)
        toolbar.addWidget(self.page_spinner)

        toolbar.addSeparator()
        self._zoom_in_action = QtGui.QAction("Zoom In", self)
        self._zoom_in_action.setShortcut(QtGui.QKeySequence.ZoomIn)
        self._zoom_in_action.triggered.connect(self.pdf_viewer.zoom_in)
        # Improved icon contrast with Fluent icons
        self._zoom_in_action.setIcon(get_fluent_icon("ZOOM_IN", "ADD", default=FI.ADD))

        self._zoom_out_action = QtGui.QAction("Zoom Out", self)
        self._zoom_out_action.setShortcut(QtGui.QKeySequence.ZoomOut)
        self._zoom_out_action.triggered.connect(self.pdf_viewer.zoom_out)
        # Improved icon contrast with Fluent icons
        self._zoom_out_action.setIcon(get_fluent_icon("ZOOM_OUT", "REMOVE", "SUBTRACT", default=FI.DELETE))

        self._zoom_fit_action = QtGui.QAction("Fit Width", self)
        self._zoom_fit_action.triggered.connect(self.pdf_viewer.fit_to_width)
        # Improved icon contrast with Fluent icons
        self._zoom_fit_action.setIcon(get_fluent_icon("FIT_PAGE", "SCALE_FILL", "FULL_SCREEN", default=FI.SYNC))

        self._zoom_actual_action = QtGui.QAction("Actual Size", self)
        self._zoom_actual_action.triggered.connect(self.pdf_viewer.actual_size)
        # Improved icon contrast with Fluent icons
        self._zoom_actual_action.setIcon(get_fluent_icon("ZOOM", "SCALE", "ZOOM_OUT", default=FI.INFO))

        self._zoom_actions = [
            self._zoom_out_action,
            self._zoom_in_action,
            self._zoom_fit_action,
            self._zoom_actual_action,
        ]

        toolbar.addActions(
            [
                self._zoom_out_action,
                self._zoom_in_action,
                self._zoom_fit_action,
                self._zoom_actual_action,
            ]
        )

        for action in (*self._zoom_actions,):
            self.addAction(action)

        self._set_output_mode(editable=True, announce=False)

        self.addToolBar(toolbar)
        self._toolbar = toolbar
        self._update_zoom_action_state()
        self._update_data_actions()

    def _create_menus(self) -> None:
        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu("&File")
        file_menu.addAction(self._import_data_action)
        file_menu.addAction(self._import_pdf_action)
        file_menu.addSeparator()
        file_menu.addAction(self._save_mapping_action)
        file_menu.addAction(self._load_mapping_action)
        file_menu.addAction(self._adjust_range_action)
        file_menu.addAction(self._edit_mapping_action)
        file_menu.addAction(self._remove_mapping_action)
        file_menu.addSeparator()
        file_menu.addAction(self._generate_action)
        file_menu.addSeparator()
        exit_action = QtGui.QAction("Exit", self)
        exit_action.setShortcut(QtGui.QKeySequence.Quit)
        # Improved icon contrast with Fluent icons
        exit_action.setIcon(get_fluent_icon("DISMISS", "CLOSE", default=FI.DELETE))
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        self._exit_action = exit_action

        view_menu = menu_bar.addMenu("&View")
        for action in self._zoom_actions:
            view_menu.addAction(action)
        view_menu.addSeparator()
        view_menu.addAction(self._toolbar.toggleViewAction())
        view_menu.addAction(self._mapping_dock.toggleViewAction())
        theme_menu = view_menu.addMenu("Theme")
        self._create_theme_actions(theme_menu)

        help_menu = menu_bar.addMenu("&Help")
        about_action = QtGui.QAction("About", self)
        # Improved icon contrast with Fluent icons
        about_action.setIcon(get_fluent_icon("INFO"))
        about_action.triggered.connect(self._show_about_dialog)
        help_menu.addAction(about_action)
        self._about_action = about_action

    # ----- Event handlers -------------------------------------------------------
    def _action_import_data(self) -> None:
        filters = ";;".join(self._data_loader.supported_filters())
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Data File", "", filters)
        if not path:
            self._set_status("Data import cancelled", timeout=3000)
            return
        self._action_import_data_from_path(Path(path))

    def _action_import_data_from_path(self, path: Path) -> None:
        try:
            sample = self._data_loader.load(path)
        except Exception as exc:  # noqa: BLE001 - display to users
            QtWidgets.QMessageBox.critical(self, "Data Import Failed", str(exc))
            self._set_status("Failed to import data source", timeout=6000)
            return

        if sample.available_sheets:
            sample = self._maybe_select_excel_sheet(path, sample)

        sample = self._maybe_prompt_data_range(path, sample)

        self._state.data_sample = sample
        self._state.mapping.source_data = sample.source_path
        self._state.mapping.data_sheet = sample.sheet_name
        self._state.mapping.header_row = sample.header_row
        self._state.mapping.data_row = sample.data_row
        self._state.mapping.column_offset = sample.column_offset
        self.spreadsheet_panel.set_data(sample)
        row_count = sample.dataframe.shape[0]
        sheet_msg = f" (sheet '{sample.sheet_name}')" if sample.sheet_name else ""
        self._set_status(f"Loaded data '{sample.source_path.name}'{sheet_msg} ({row_count:,} rows)")
        self._refresh_mapping_labels()
        self._update_data_actions()
        self._update_preview_data_indicator()

    def _action_adjust_data_range(self) -> None:
        sample = self._state.data_sample
        source = self._state.mapping.source_data
        if not sample or not source:
            QtWidgets.QMessageBox.information(self, "No Data", "Load a data file before adjusting the range.")
            return
        adjusted = self._prompt_data_range(source, sample)
        if adjusted is sample:
            return
        self._state.data_sample = adjusted
        self._state.mapping.source_data = adjusted.source_path
        self._state.mapping.data_sheet = adjusted.sheet_name
        self._state.mapping.header_row = adjusted.header_row
        self._state.mapping.data_row = adjusted.data_row
        self._state.mapping.column_offset = adjusted.column_offset
        self.spreadsheet_panel.set_data(adjusted)
        row_count = adjusted.dataframe.shape[0]
        sheet_msg = f" (sheet '{adjusted.sheet_name}')" if adjusted.sheet_name else ""
        self._set_status(
            f"Adjusted data range for '{adjusted.source_path.name}'{sheet_msg} ({row_count:,} rows)", timeout=6000
        )
        self._refresh_mapping_labels()
        self._update_data_actions()
        self._update_preview_data_indicator()

    def _action_import_pdf(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select PDF Template", "", "PDF Files (*.pdf)"
        )
        if not path:
            self._set_status("PDF import cancelled", timeout=3000)
            return
        self._action_import_pdf_from_path(Path(path))

    def _action_import_pdf_from_path(self, path: Path) -> None:
        try:
            template = self._pdf_engine.open_template(path)
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "PDF Import Failed", str(exc))
            self._show_pdf_placeholder()
            self._set_status("Failed to import PDF template", timeout=6000)
            return

        if self._state.pdf_template:
            self._state.pdf_template.close()
        self._state.pdf_template = template
        self._state.mapping.pdf_template = template.path
        self.pdf_viewer.load_template(template)
        self.pdf_viewer.actual_size()
        self._show_pdf_viewer()
        page_count = template.document.page_count
        self._set_status(f"Loaded PDF '{template.path.name}' ({page_count} pages)")
        self._refresh_mapping_labels()
        self._update_data_actions()
        self._configure_page_controls_for_template()

    def _action_save_mapping(self) -> None:
        if not self._state.mapping.rules:
            QtWidgets.QMessageBox.information(self, "No Mappings", "Create at least one mapping first.")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Mapping", "", "Mapping Files (*.json)"
        )
        if not path:
            self._set_status("Save mapping cancelled", timeout=3000)
            return
        try:
            self._mapping_manager.save(Path(path), self._state.mapping)
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Save Failed", str(exc))
            self._set_status("Failed to save mapping", timeout=6000)
        else:
            self._set_status(f"Saved mapping to '{Path(path).name}'", timeout=6000)

    def _action_load_mapping(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load Mapping", "", "Mapping Files (*.json)"
        )
        if not path:
            self._set_status("Load mapping cancelled", timeout=3000)
            return
        self._action_load_mapping_from_path(Path(path))

    def _action_load_mapping_from_path(self, path: Path) -> None:
        try:
            mapping = self._mapping_manager.load(path)
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Load Failed", str(exc))
            self._set_status("Failed to load mapping file", timeout=6000)
            return

        self._state.mapping = mapping
        if mapping.source_data and mapping.source_data.exists():
            try:
                sample = self._data_loader.load(
                    mapping.source_data,
                    sheet=mapping.data_sheet,
                    header_row=mapping.header_row,
                    data_row=mapping.data_row,
                    column_offset=(mapping.column_offset or 0),
                )
            except Exception:  # noqa: BLE001
                sample = None
            else:
                if sample.available_sheets and mapping.data_sheet not in sample.available_sheets:
                    sample = self._maybe_select_excel_sheet(mapping.source_data, sample)
                self._state.data_sample = sample
                self._state.mapping.source_data = sample.source_path
                self._state.mapping.data_sheet = sample.sheet_name
                self._state.mapping.header_row = sample.header_row
                self._state.mapping.data_row = sample.data_row
                self._state.mapping.column_offset = sample.column_offset
                self.spreadsheet_panel.set_data(sample)
                self.pdf_viewer.clear_field_selection()
                self._update_preview_data_indicator()
                sheet_msg = f" (sheet '{sample.sheet_name}')" if sample.sheet_name else ""
                self._set_status(
                    f"Loaded data '{sample.source_path.name}'{sheet_msg} ({sample.dataframe.shape[0]:,} rows)"
                )
        else:
            self.spreadsheet_panel.clear()
            self.pdf_viewer.clear_field_selection()
            self._update_preview_data_indicator()
            self._state.data_sample = None
            self._state.mapping.source_data = None
            self._state.mapping.data_sheet = None
            self._state.mapping.header_row = None
            self._state.mapping.data_row = None
            self._state.mapping.column_offset = None
            self._set_status("Mapping references missing data source", timeout=5000)

        if mapping.pdf_template and mapping.pdf_template.exists():
            try:
                template = self._pdf_engine.open_template(mapping.pdf_template)
            except Exception:  # noqa: BLE001
                template = None
            else:
                if self._state.pdf_template:
                    self._state.pdf_template.close()
                self._state.pdf_template = template
                self.pdf_viewer.load_template(template)
                self.pdf_viewer.actual_size()
                self._show_pdf_viewer()
                self._set_status(
                    f"Loaded PDF '{template.path.name}' ({template.document.page_count} pages)"
                )
        else:
            if self._state.pdf_template:
                self._state.pdf_template.close()
            self._state.pdf_template = None
            self.pdf_viewer.clear()
            self._show_pdf_placeholder()

        self._refresh_mapping_labels()
        self._configure_page_controls_for_template()
        self._set_status(f"Loaded mapping '{Path(path).name}'", timeout=6000)
        self._update_data_actions()
        self._update_preview_data_indicator()

    def _load_generation_defaults(self) -> Dict[str, Any]:
        settings = self._settings
        mode_value = settings.value(GENERATION_MODE_KEY, "per_entry")
        directory = self._read_path_setting(GENERATION_DIR_KEY)
        file_path = self._read_path_setting(GENERATION_FILE_KEY)
        columns = self._read_list_setting(GENERATION_COLUMNS_KEY)
        prefix_value = settings.value(GENERATION_PREFIX_KEY, "")
        suffix_value = settings.value(GENERATION_SUFFIX_KEY, "")
        separator_value = settings.value(GENERATION_SEPARATOR_KEY, "_")
        read_only_value = settings.value(GENERATION_READ_ONLY_KEY, False, type=bool)
        return {
            "mode": str(mode_value) if mode_value else "per_entry",
            "directory": directory,
            "file": file_path,
            "columns": columns,
            "prefix": str(prefix_value) if prefix_value else "",
            "suffix": str(suffix_value) if suffix_value else "",
            "separator": str(separator_value) if separator_value else "_",
            "read_only": bool(read_only_value),
        }

    def _persist_generation_defaults(self, options: GenerationOptions) -> None:
        settings = self._settings
        settings.setValue(GENERATION_MODE_KEY, options.mode)
        if options.destination_dir:
            settings.setValue(GENERATION_DIR_KEY, str(options.destination_dir))
        if options.combined_path:
            settings.setValue(GENERATION_FILE_KEY, str(options.combined_path))
        settings.setValue(GENERATION_COLUMNS_KEY, "|".join(options.columns))
        settings.setValue(GENERATION_PREFIX_KEY, options.prefix)
        settings.setValue(GENERATION_SUFFIX_KEY, options.suffix)
        settings.setValue(GENERATION_SEPARATOR_KEY, options.separator)
        settings.setValue(GENERATION_READ_ONLY_KEY, options.read_only)

    def _read_path_setting(self, key: str) -> Path | None:
        value = self._settings.value(key)
        if not value:
            return None
        try:
            return Path(str(value))
        except TypeError:
            return None

    def _read_list_setting(self, key: str) -> list[str]:
        value = self._settings.value(key, [])
        if isinstance(value, str):
            if not value:
                return []
            return [item for item in value.split("|") if item]
        if isinstance(value, (list, tuple)):
            return [str(item) for item in value if item]
        return []
    def _action_generate_pdfs(self) -> None:
        if not self._state.data_sample or not self._state.pdf_template:
            QtWidgets.QMessageBox.warning(
                self, "Missing Data", "Load both a data file and a PDF template first."
            )
            self._set_status("Cannot generate PDFs without data and template", timeout=5000)
            return
        if not self._state.mapping.rules:
            QtWidgets.QMessageBox.warning(self, "Missing Mapping", "Map at least one field first.")
            self._set_status("Create at least one mapping before generating PDFs", timeout=5000)
            return

        if self._generation_thread and self._generation_thread.isRunning():
            QtWidgets.QMessageBox.information(
                self, "Generation Running", "A generation task is already in progress."
            )
            self._set_status("PDF generation already running", timeout=4000)
            return

        rows = self._state.data_sample.dataframe.to_dict(orient="records")
        if not rows:
            QtWidgets.QMessageBox.information(self, "No Rows", "The selected dataset is empty.")
            self._set_status("Data source contains no rows", timeout=4000)
            return

        defaults = self._load_generation_defaults()
        columns = list(self._state.data_sample.dataframe.columns)
        sample_row = rows[0] if rows else {}

        dialog = GeneratePdfDialog(
            self,
            columns=columns,
            sample_row=sample_row,
            default_mode=defaults["mode"],
            default_directory=defaults["directory"],
            default_file=defaults["file"],
            default_columns=defaults["columns"],
            default_prefix=defaults["prefix"],
            default_suffix=defaults["suffix"],
            default_separator=defaults["separator"],
            default_read_only=defaults["read_only"],
        )
        if dialog.exec() != QtWidgets.QDialog.Accepted:
            self._set_status("PDF generation cancelled", timeout=4000)
            return

        options = dialog.options()
        self._persist_generation_defaults(options)

        builder = FilenameBuilder(
            options.columns,
            prefix=options.prefix,
            suffix=options.suffix,
            separator=options.separator or "_",
        )

        read_only_choice = options.read_only

        self._set_output_mode(editable=not read_only_choice, announce=False)

        output_dir: Path | None = None
        combined_path: Path | None = None
        try:
            if options.mode == "per_entry":
                if options.destination_dir is None:
                    raise ValueError("Select an output folder.")
                output_dir = options.destination_dir.expanduser().resolve()
                output_dir.mkdir(parents=True, exist_ok=True)
            else:
                if options.combined_path is None:
                    raise ValueError("Choose a combined PDF filename.")
                combined_path = options.combined_path.expanduser().resolve()
                combined_path = combined_path.with_suffix(".pdf")
                combined_path.parent.mkdir(parents=True, exist_ok=True)
                output_dir = combined_path.parent
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Destination Error", str(exc))
            return

        behavior = "read-only" if read_only_choice else "editable"
        if options.mode == "per_entry":
            status_target = output_dir.name if output_dir else ""
            self._set_status(
                f"Generating {len(rows)} PDFs into '{status_target}' ({behavior})",
                timeout=5000,
            )
        else:
            assert combined_path is not None
            self._set_status(
                f"Generating combined PDF '{combined_path.name}' ({behavior})", timeout=5000
            )

        progress = QtWidgets.QProgressDialog(
            "Generating PDFs...", "Cancel", 0, len(rows), self, QtCore.Qt.WindowTitleHint
        )
        progress.setWindowModality(QtCore.Qt.WindowModal)
        progress.setValue(0)

        flatten_output = False

        worker = PdfGenerationWorker(
            self._pdf_engine,
            self._state.pdf_template.path,
            output_dir or Path.cwd(),
            list(self._state.mapping.iter_rules()),
            rows,
            flatten=flatten_output,
            read_only=read_only_choice,
            template_metadata=self._state.pdf_template,
            mode=options.mode,
            combined_output=combined_path,
            filename_builder=builder,
        )

        thread = QtCore.QThread(self)
        worker.moveToThread(thread)

        worker.progress.connect(self._on_generation_progress)
        worker.completed.connect(self._on_generation_completed)
        worker.failed.connect(self._on_generation_failed)
        worker.cancelled.connect(self._on_generation_cancelled)

        thread.started.connect(worker.run)
        thread.finished.connect(thread.deleteLater)
        progress.canceled.connect(worker.request_cancel)

        self._generation_thread = thread
        self._generation_worker = worker
        self._generation_progress = progress
        self._last_generation_mode = options.mode
        self._last_generation_target = combined_path if combined_path else output_dir
        self._last_generation_read_only = read_only_choice

        thread.start()
        progress.show()

    # ----- Helpers -------------------------------------------------------------
    def _on_field_assigned(self, field_name: str, column_name: str) -> None:
        self._state.mapping.assign(field_name, column_name)
        self._refresh_mapping_labels()
        self.pdf_viewer.clear_field_selection()
        self._update_preview_data_indicator()

    def _on_field_remove_requested(self, field_name: str) -> None:
        self._action_remove_mapping(field_name)

    def _refresh_mapping_labels(self) -> None:
        rules = self._state.mapping.rules
        sample_row = self._current_sample_row()
        sample_payload: Dict[str, str] = {}
        if sample_row:
            sample_payload = evaluate_rules(rules.values(), sample_row)

        previews: Dict[str, str] = {}
        for field_name, rule in rules.items():
            previews[field_name] = self._format_rule_preview(rule, sample_payload)

        self.mapping_table.update_mapping(rules, previews)
        self.pdf_viewer.clear_assignments()
        for _, rule in sorted(rules.items(), key=lambda item: item[0]):
            descriptor = rule.describe()
            for target in rule.targets:
                raw_value = sample_payload.get(target, "")
                preview_value = self._render_preview_value(raw_value)
                self.pdf_viewer.set_assignment(target, descriptor, preview_value)
        self._update_mapping_action_state()

    def _current_sample_row(self) -> Optional[Dict[str, object]]:
        if not self._state.data_sample:
            return None
        frame = self._state.data_sample.dataframe
        if frame.empty:
            return None
        return frame.iloc[0].to_dict()

    def _format_rule_preview(self, rule: MappingRule, payload: Dict[str, str]) -> str:
        if not payload:
            return ""
        values = []
        for target in rule.targets:
            raw_value = payload.get(target, "")
            display_value = self._render_preview_value(raw_value)
            if display_value:
                values.append(f"{target}: {display_value}")
        if values:
            if len(rule.targets) == 1:
                return values[0].split(": ", 1)[-1]
            return "; ".join(values)
        for target in rule.targets:
            if target in payload:
                return self._render_preview_value(payload.get(target, ""))
        return ""

    def _render_preview_value(self, value: object) -> str:
        """Return the text to show in previews, mimicking checkbox appearance."""
        checkmark = "\u2713"
        kind, normalized = PdfEngine._normalize_payload_value(value)
        if kind != "checkbox":
            return str(normalized)
        if isinstance(normalized, str):
            stripped = normalized.strip()
            lowered = stripped.lstrip("/").lower()
            if lowered in {"", "off", "no", "false", "0", "unchecked"}:
                return ""
            return checkmark
        if isinstance(normalized, bool):
            return checkmark if normalized else ""
        return checkmark if normalized else ""

    def _show_pdf_viewer(self) -> None:
        self.viewer_stack.setCurrentWidget(self.pdf_viewer)
        self._update_zoom_action_state()
        self._on_zoom_changed(self.pdf_viewer.current_zoom())

    def _show_pdf_placeholder(self) -> None:
        self.viewer_stack.setCurrentWidget(self._pdf_placeholder)
        self._update_zoom_action_state()

    def _maybe_select_excel_sheet(self, path: Path, sample: DataSample) -> DataSample:
        sheets = list(sample.available_sheets)
        if len(sheets) <= 1:
            return sample
        current_sheet = sample.sheet_name or self._state.mapping.data_sheet
        try:
            default_index = sheets.index(current_sheet) if current_sheet in sheets else 0
        except ValueError:
            default_index = 0

        sheet, ok = QtWidgets.QInputDialog.getItem(
            self,
            "Select Worksheet",
            f"Select worksheet from {path.name}:",
            sheets,
            default_index,
            False,
        )
        if not ok or not sheet:
            self._set_status("Worksheet selection cancelled", timeout=4000)
            return sample
        if sheet == sample.sheet_name:
            return sample
        try:
            return self._data_loader.load(path, sheet=sheet)
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Worksheet Load Failed", str(exc))
            self._set_status(f"Failed to load worksheet '{sheet}'", timeout=6000)
            return sample

    def _prompt_data_range(self, path: Path, sample: DataSample) -> DataSample:
        dialog = DataRangeDialog(
            self,
            header_row=sample.header_row or 1,
            data_row=sample.data_row or max((sample.header_row or 1) + 1, 2),
            first_column=(sample.column_offset or 0) + 1,
        )
        if dialog.exec() != QtWidgets.QDialog.Accepted:
            return sample
        header_row, data_row, first_column = dialog.values()
        try:
            return self._data_loader.load(
                path,
                sheet=sample.sheet_name,
                header_row=header_row,
                data_row=data_row,
                column_offset=max(0, first_column - 1),
            )
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Data Range Invalid", str(exc))
            self._set_status("Failed to load data with provided offsets", timeout=6000)
            return sample

    def _maybe_prompt_data_range(self, path: Path, sample: DataSample) -> DataSample:
        prompt = QtWidgets.QMessageBox.question(
            self,
            "Adjust Data Range?",
            "Do column headers or data rows start lower in this file?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )
        if prompt == QtWidgets.QMessageBox.Yes:
            return self._prompt_data_range(path, sample)
        return sample

    def _on_page_selected(self, value: int) -> None:
        if not self._state.pdf_template:
            return
        self.pdf_viewer.set_page(value - 1)

    def _on_page_changed(self, page_index: int) -> None:
        if not hasattr(self, "page_spinner"):
            return
        self.page_spinner.blockSignals(True)
        self.page_spinner.setValue(page_index + 1)
        self.page_spinner.blockSignals(False)
        self._configure_page_controls_for_template()
        self._refresh_mapping_labels()

    def _on_zoom_changed(self, zoom: float) -> None:
        if hasattr(self, "_zoom_label"):
            self._zoom_label.setText(f"{int(round(zoom * 100))}%")

    def _update_mapping_action_state(self) -> None:
        selected = self.mapping_table.selected_field()
        if hasattr(self, "_edit_mapping_action"):
            self._edit_mapping_action.setEnabled(selected is not None)
        if hasattr(self, "_remove_mapping_action"):
            self._remove_mapping_action.setEnabled(selected is not None)
        self._update_data_actions()
        self._update_zoom_action_state()

    def _update_zoom_action_state(self) -> None:
        enabled = self._state.pdf_template is not None
        for action in getattr(self, "_zoom_actions", []):
            action.setEnabled(enabled)
        if hasattr(self, "_zoom_label"):
            self._zoom_label.setEnabled(enabled)
            if not enabled:
                self._zoom_label.setText("--")

    def _update_data_actions(self) -> None:
        if hasattr(self, "_adjust_range_action"):
            self._adjust_range_action.setEnabled(self._state.data_sample is not None)

    def _set_output_mode(self, *, editable: bool, announce: bool = False) -> None:
        self._read_only_output = not editable

        self._update_output_mode_label(editable)

        if announce:
            mode_text = (
                "Output mode set to: Editable (fillable PDFs)"
                if editable
                else "Output mode set to: Read-Only (fields locked)"
            )
            self._set_status(mode_text, timeout=5000)

    def _update_output_mode_label(self, editable: bool) -> None:
        if not hasattr(self, "_output_mode_label"):
            return
        text = (
            "Output Mode: Editable (fillable)"
            if editable
            else "Output Mode: Read-Only (locked)"
        )
        self._output_mode_label.setText(text)
        palette = self._output_mode_label.palette()
        palette.setColor(
            QtGui.QPalette.WindowText,
            QtGui.QColor("#2f7d3b") if editable else QtGui.QColor("#a83232"),
        )
        self._output_mode_label.setPalette(palette)

    def _update_preview_data_indicator(self) -> None:
        """Refresh the status indicator describing the previewed dataset row."""
        label = getattr(self, "_preview_data_label", None)
        if label is None:
            return

        sample = self._state.data_sample
        if sample is None:
            label.setText("Preview data: none loaded")
            label.setToolTip("Load a dataset to drive the live PDF preview.")
            return

        rows = sample.dataframe.shape[0]
        if rows == 0:
            label.setText("Preview data: dataset empty")
            label.setToolTip("The loaded dataset contains no rows to preview.")
            return

        sheet = f" - Sheet: {sample.sheet_name}" if sample.sheet_name else ""
        selected_field = self._selected_viewer_field
        base_text = "Previewing data row 1"
        if selected_field:
            base_text = f"{base_text} - Field selected: {selected_field}"
        label.setText(base_text)
        tooltip = f"Showing values from the first row of the dataset{sheet}."
        if selected_field:
            tooltip += f" Field '{selected_field}' is ready for column assignment."
        label.setToolTip(tooltip)

    def _on_field_selection_changed(self, field_name: Optional[str]) -> None:
        self._selected_viewer_field = field_name or None
        if field_name is None:
            self.spreadsheet_panel.columns_widget.clearSelection()
        self._update_preview_data_indicator()

    def _on_column_activated(self, column_name: str) -> None:
        if not self._selected_viewer_field:
            return
        field_name = self._selected_viewer_field
        self._on_field_assigned(field_name, column_name)
        self._set_status(f"Assigned column '{column_name}' to '{field_name}'", timeout=3000)

    def _show_about_dialog(self) -> None:
        QtWidgets.QMessageBox.about(
            self,
            "About PDF Bulk Filler",
            (
                "<b>PDF Bulk Filler</b><br>"
                "Version 0.1.0<br><br>"
                "Map tabular data onto fillable PDF forms with a drag-and-drop workflow. "
                "All processing happens locally so your data stays private."
            ),
        )

    def _action_edit_mapping(self, field_name: Optional[str] = None) -> None:
        field = field_name or self.mapping_table.selected_field()
        if not field:
            return
        if not self._state.pdf_template:
            QtWidgets.QMessageBox.warning(
                self,
                "No PDF Template",
                "Load a PDF template before editing mapping rules.",
            )
            return
        available_fields = sorted({f.field_name for f in self._state.pdf_template.fields})
        available_columns: list[str] = []
        if self._state.data_sample is not None:
            available_columns = list(self._state.data_sample.columns())
        rule = self._state.mapping.resolve(field)
        remove_callback = None
        if field in self._state.mapping.rules:
            remove_callback = lambda f=field: self._action_remove_mapping(f)
        dialog = RuleEditorDialog(
            field,
            rule,
            available_fields,
            available_columns,
            self,
            remove_callback=remove_callback,
        )
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            updated_rule = dialog.selected_rule()
            self._state.mapping.assign(field, updated_rule)
            self._refresh_mapping_labels()
            self._set_status(f"Updated mapping rule for '{field}'", timeout=4000)

    def _action_remove_mapping(self, field_name: Optional[str] = None) -> None:
        field = field_name or self.mapping_table.selected_field()
        if not field:
            return
        if field in self._state.mapping.rules:
            self._state.mapping.remove(field)
            self._refresh_mapping_labels()
            self.pdf_viewer.clear_field_selection()
            self._update_data_actions()
            self._update_preview_data_indicator()
            self._set_status(f"Removed mapping for '{field}'", timeout=4000)

    def _configure_page_controls_for_template(self) -> None:
        if not hasattr(self, "page_spinner"):
            return
        if self._state.pdf_template:
            self.page_spinner.blockSignals(True)
            self.page_spinner.setMaximum(
                max(1, self._state.pdf_template.document.page_count)
            )
            self.page_spinner.setValue(self.pdf_viewer.current_page() + 1)
            self.page_spinner.blockSignals(False)
            self.page_spinner.setEnabled(True)
            self._show_pdf_viewer()
        else:
            self.page_spinner.blockSignals(True)
            self.page_spinner.setValue(1)
            self.page_spinner.blockSignals(False)
            self.page_spinner.setEnabled(False)
            self._show_pdf_placeholder()

    def _on_generation_progress(self, current: int, total: int) -> None:
        if not self._generation_progress:
            return
        self._generation_progress.setMaximum(total)
        self._generation_progress.setValue(current)
        self._set_status(f"Generating PDFs... {current}/{total}", timeout=1500)

    def _on_generation_completed(self, outputs: list[Path]) -> None:
        mode = getattr(self, "_last_generation_mode", "per_entry")
        target = getattr(self, "_last_generation_target", None)
        self._cleanup_generation_worker()

        if mode == "combined":
            final_path = outputs[0] if outputs else target
            if final_path:
                final_path = Path(final_path)
                QtWidgets.QMessageBox.information(
                    self,
                    "Generation Complete",
                    f"Created combined PDF at {final_path}.",
                )
                self._set_status(f"Combined PDF saved as {final_path.name}", timeout=6000)
                self._last_generation_target = None
            else:
                self._set_status("Combined PDF created", timeout=6000)
                self._last_generation_target = None
            return

        destination = target or (outputs[0].parent if outputs else None)
        if destination:
            QtWidgets.QMessageBox.information(
                self,
                "Generation Complete",
                f"Created {len(outputs)} PDF files in {destination}.",
            )
            self._set_status(f"Created {len(outputs)} PDF files", timeout=6000)
            self._last_generation_target = None
        else:
            self._last_generation_target = None

    def _on_generation_failed(self, message: str) -> None:
        self._cleanup_generation_worker()
        self._last_generation_target = None
        QtWidgets.QMessageBox.critical(self, "Generation Failed", message)
        self._set_status("PDF generation failed", timeout=6000)

    def _on_generation_cancelled(self) -> None:
        self._cleanup_generation_worker()
        self._last_generation_target = None
        QtWidgets.QMessageBox.information(self, "Generation Cancelled", "PDF creation cancelled.")
        self._set_status("PDF generation cancelled", timeout=4000)

    def _cleanup_generation_worker(self) -> None:
        if self._generation_progress:
            self._generation_progress.close()
            self._generation_progress = None
        if self._generation_worker:
            self._generation_worker.deleteLater()
            self._generation_worker = None
        if self._generation_thread:
            self._generation_thread.quit()
            self._generation_thread.wait()
            self._generation_thread = None

    def _set_status(self, message: str, timeout: int = 4000) -> None:
        self.statusBar().showMessage(message, timeout)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # noqa: N802
        if self._generation_worker:
            self._generation_worker.request_cancel()
        self._cleanup_generation_worker()
        if self._state.pdf_template:
            self._state.pdf_template.close()
        super().closeEvent(event)

    # ----- Theme management ----------------------------------------------------
    def _create_theme_actions(self, menu: QtWidgets.QMenu) -> None:
        # Added theme switcher and persistence
        action_group = QtGui.QActionGroup(self)
        action_group.setExclusive(True)
        self._theme_actions: dict[str, QtGui.QAction] = {}

        for mode, label in THEME_LABELS.items():
            action = QtGui.QAction(label, self)
            action.setCheckable(True)
            action_group.addAction(action)
            action.triggered.connect(lambda checked, m=mode: checked and self._apply_theme(m))
            menu.addAction(action)
            self._theme_actions[mode] = action

        self._update_theme_action_checks()

    def _apply_theme(self, mode: str, *, save: bool = True, update_actions: bool = True) -> None:
        # Added theme switcher and persistence
        if mode not in THEME_MAP:
            mode = DEFAULT_THEME_MODE
        setTheme(THEME_MAP[mode])
        setThemeColor(ACCENT_COLOR)
        # Added theme switcher and persistence
        self._apply_palette_for_theme(qconfig.theme)
        self._theme_mode = mode
        if save:
            self._settings.setValue(THEME_SETTINGS_KEY, mode)
        if update_actions:
            self._update_theme_action_checks()

    def _update_theme_action_checks(self) -> None:
        if not hasattr(self, "_theme_actions"):
            return
        for mode, action in self._theme_actions.items():
            action.setChecked(mode == self._theme_mode)

    def _apply_palette_for_theme(self, theme: Theme) -> None:
        """Synchronize the Qt palette with the active Fluent theme."""
        app = QtWidgets.QApplication.instance()
        if app is None:
            return

        primary = ThemeColor.PRIMARY.color()
        app.setStyle("Fusion")

        if theme == Theme.DARK:
            palette = QtGui.QPalette()
            base_color = QtGui.QColor("#202020")
            alt_base = QtGui.QColor("#2a2a2a")
            text_color = QtGui.QColor("#f0f0f0")
            disabled_text = QtGui.QColor("#8c8c8c")

            palette.setColor(QtGui.QPalette.Window, base_color)
            palette.setColor(QtGui.QPalette.Base, QtGui.QColor("#1a1a1a"))
            palette.setColor(QtGui.QPalette.AlternateBase, alt_base)
            palette.setColor(QtGui.QPalette.ToolTipBase, alt_base)
            palette.setColor(QtGui.QPalette.ToolTipText, text_color)
            palette.setColor(QtGui.QPalette.Text, text_color)
            palette.setColor(QtGui.QPalette.Button, base_color)
            palette.setColor(QtGui.QPalette.ButtonText, text_color)
            palette.setColor(QtGui.QPalette.WindowText, text_color)
            palette.setColor(QtGui.QPalette.Highlight, primary)
            palette.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor("#ffffff"))
            palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, disabled_text)
            palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, disabled_text)
            palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, disabled_text)
            palette.setColor(QtGui.QPalette.Link, primary)
            palette.setColor(QtGui.QPalette.LinkVisited, primary.darker(110))
        else:
            palette = QtGui.QPalette()
            background = QtGui.QColor("#f4f4f6")
            base = QtGui.QColor("#ffffff")
            alt_base = QtGui.QColor("#f1f1f1")
            text_color = QtGui.QColor("#1f1f23")
            disabled_text = QtGui.QColor("#9b9b9f")

            palette.setColor(QtGui.QPalette.Window, background)
            palette.setColor(QtGui.QPalette.Base, base)
            palette.setColor(QtGui.QPalette.AlternateBase, alt_base)
            palette.setColor(QtGui.QPalette.ToolTipBase, base)
            palette.setColor(QtGui.QPalette.ToolTipText, text_color)
            palette.setColor(QtGui.QPalette.Text, text_color)
            palette.setColor(QtGui.QPalette.Button, QtGui.QColor("#efeff1"))
            palette.setColor(QtGui.QPalette.ButtonText, text_color)
            palette.setColor(QtGui.QPalette.WindowText, text_color)
            palette.setColor(QtGui.QPalette.Mid, QtGui.QColor("#d3d3d9"))
            palette.setColor(QtGui.QPalette.Light, QtGui.QColor("#ffffff"))
            palette.setColor(QtGui.QPalette.Dark, QtGui.QColor("#b8b8bf"))
            palette.setColor(QtGui.QPalette.Shadow, QtGui.QColor("#a8a8af"))
            palette.setColor(QtGui.QPalette.Highlight, primary)
            palette.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor("#ffffff"))
            palette.setColor(QtGui.QPalette.Link, primary.darker(110))
            palette.setColor(QtGui.QPalette.LinkVisited, primary.darker(130))
            palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, disabled_text)
            palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, disabled_text)
            palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, disabled_text)

        app.setPalette(palette)
        self._apply_tooltip_style(theme)
        self._repolish_for_theme()

    def _repolish_for_theme(self) -> None:
        """Force widgets to re-polish so palette changes take effect."""
        widgets = [
            self,
            self.centralWidget(),
            getattr(self, "spreadsheet_panel", None),
            getattr(self, "viewer_stack", None),
            getattr(self, "mapping_table", None),
        ]
        for widget in widgets:
            if widget is None:
                continue
            widget.style().unpolish(widget)
            widget.style().polish(widget)
            if hasattr(widget, "viewport"):
                widget.viewport().update()
            widget.repaint()

    def _apply_tooltip_style(self, theme: Theme) -> None:
        if not hasattr(self, "_tooltip_styler") or self._tooltip_styler is None:
            self._tooltip_styler = TooltipStyler(self)

        if theme == Theme.DARK:
            background = "#262b34"
            text = "#f5f5f8"
            shadow_color = QtGui.QColor(0, 0, 0, 160)
        else:
            background = "#ffffff"
            text = "#1f1f23"
            shadow_color = QtGui.QColor(31, 31, 35, 100)

        border_color = QtGui.QColor(background)
        border_color.setAlphaF(0.8)

        tooltip_stylesheet = (
            "QToolTip {"
            f" color: {text};"
            f" background-color: {background};"
            f" border: 1px solid {border_color.name(QtGui.QColor.HexArgb)};"
            " border-radius: 2px;"
            " padding: 2px 4px;"
            "}"
        )

        app = QtWidgets.QApplication.instance()
        if app is None:
            return

        existing = app.styleSheet() or ""
        cleaned = re.sub(r"QToolTip\s*\{[^}]*\}", "", existing, flags=re.DOTALL).strip()
        if cleaned:
            cleaned = f"{cleaned}\n"
        app.setStyleSheet(f"{cleaned}{tooltip_stylesheet}")

        self._tooltip_styler.update_theme(
            background=background,
            text=text,
            border=border_color,
            shadow=shadow_color,
        )
        if hasattr(self, '_tooltip_manager') and self._tooltip_manager is not None:
            self._tooltip_manager.update_theme(
                background=background,
                text=text,
                border=border_color,
                shadow=shadow_color,
            )
























