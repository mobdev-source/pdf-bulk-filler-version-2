"""PySide6 main window containing the drag-and-drop mapping interface."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, Iterable, Optional

import pandas as pd
import numpy as np
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

from pdf_bulk_filler.data.loader import DataLoader, DataSample
from pdf_bulk_filler.mapping.manager import MappingManager, MappingModel
from pdf_bulk_filler.pdf.engine import PdfEngine, PdfField, PdfTemplate
from pdf_bulk_filler.ui.workers import PdfGenerationWorker


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

    def __init__(self) -> None:
        super().__init__()
        self.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.setDragEnabled(True)
        self.setAlternatingRowColors(True)

    def set_columns(self, columns: Iterable[str]) -> None:
        self.clear()
        for column in columns:
            self.addItem(column)

    def mimeData(self, items: list[QtWidgets.QListWidgetItem]) -> QtCore.QMimeData:  # noqa: N802
        mime = QtCore.QMimeData()
        if items:
            mime.setText(items[0].text())
        return mime



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
        column_layout.addWidget(QtWidgets.QLabel("Columns"))
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

        layout.addLayout(summary_layout)
        layout.addWidget(splitter)

    def set_data(self, sample: DataSample) -> None:
        self.columns_widget.set_columns(sample.columns())
        preview = sample.head_records(50)
        self.table_model.update(preview)
        rows, cols = sample.dataframe.shape
        sheet_suffix = f" | Sheet: {sample.sheet_name}" if sample.sheet_name else ""
        self.data_summary_label.setText(f"{rows:,} rows x {cols} columns{sheet_suffix}")

    def clear(self) -> None:
        self.columns_widget.clear()
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


class PdfFieldItem(QtWidgets.QGraphicsRectItem):
    """Interactive overlay representing a PDF form field on the canvas."""

    def __init__(
        self,
        field: PdfField,
        rect: QtCore.QRectF,
        drop_callback: Callable[[PdfField, str], None],
    ) -> None:
        super().__init__(rect)
        self.field = field
        self._drop_callback = drop_callback
        self.setBrush(QtGui.QColor(0, 170, 255, 50))
        self.setPen(QtGui.QPen(QtGui.QColor(0, 120, 215), 1, QtCore.Qt.DashLine))
        self.setZValue(1)
        self.setAcceptDrops(True)
        self._label: QtWidgets.QGraphicsSimpleTextItem | None = None

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
        if column_name:
            self.setToolTip(f"PDF Field: {self.field.field_name}\nColumn: {column_name}")
            self.setBrush(QtGui.QColor(0, 170, 255, 100))
            label = self._ensure_label()
            label.setPos(self.rect().left() + 2, self.rect().top() + 2)
            label.setText((sample_value or "")[:64])
        else:
            self.setToolTip(f"PDF Field: {self.field.field_name}")
            self.setBrush(QtGui.QColor(0, 170, 255, 50))
            if self._label:
                self._label.setText("")

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


class PdfViewerWidget(QtWidgets.QGraphicsView):
    """Displays a rendered PDF page with draggable form fields."""

    fieldAssigned = QtCore.Signal(str, str)
    pageChanged = QtCore.Signal(int)
    zoomChanged = QtCore.Signal(float)

    def __init__(self, engine: PdfEngine) -> None:
        super().__init__()
        self._engine = engine
        self._template: PdfTemplate | None = None
        self._zoom = 1.0
        self._current_page = 0
        self._page_count = 0
        self._field_items: Dict[str, PdfFieldItem] = {}
        self._auto_fit = False
        self._assignments: Dict[str, tuple[str | None, str | None]] = {}

        self.setScene(QtWidgets.QGraphicsScene(self))
        self.setRenderHint(QtGui.QPainter.Antialiasing)
        self.setAcceptDrops(True)
        self.setAlignment(QtCore.Qt.AlignCenter)

    def load_template(self, template: PdfTemplate, page_index: int = 0) -> None:
        self._template = template
        self._assignments.clear()
        self._page_count = template.document.page_count
        self._current_page = max(0, min(page_index, self._page_count - 1))
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
        self.zoomChanged.emit(self._zoom)

    def _handle_drop(self, field: PdfField, column_name: str) -> None:
        self.fieldAssigned.emit(field.field_name, column_name)

    def set_assignment(
        self,
        field_name: str,
        column_name: Optional[str],
        sample_value: Optional[str] = None,
    ) -> None:
        if item := self._field_items.get(field_name):
            item.update_assignment(column_name, sample_value)
        if column_name:
            self._assignments[field_name] = (column_name, sample_value)
        else:
            self._assignments.pop(field_name, None)

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
            item = PdfFieldItem(field, rect, self._handle_drop)
            self.scene().addItem(item)
            self._field_items[field.field_name] = item

        for field_name, (column, preview) in self._assignments.items():
            if item := self._field_items.get(field_name):
                item.update_assignment(column, preview)

        self.scene().setSceneRect(self.scene().itemsBoundingRect())
        self.resetTransform()
        self.centerOn(self.scene().sceneRect().center())

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


class MappingTable(QtWidgets.QTableWidget):
    """Tabular display of current mappings."""

    removeRequested = QtCore.Signal(str)

    def __init__(self) -> None:
        super().__init__(0, 3)
        self.setHorizontalHeaderLabels(["Field", "Column", ""])
        header = self.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        self.verticalHeader().setVisible(False)
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)

    def update_mapping(self, assignments: Dict[str, str]) -> None:
        self.setRowCount(len(assignments))
        for row, (field, column) in enumerate(sorted(assignments.items())):
            self.setItem(row, 0, QtWidgets.QTableWidgetItem(field))
            self.setItem(row, 1, QtWidgets.QTableWidgetItem(column))
            remove_button = QtWidgets.QToolButton()
            # Improved icon contrast with Fluent icons
            remove_button.setIcon(get_fluent_icon("DELETE"))
            remove_button.setIconSize(QtCore.QSize(16, 16))
            remove_button.setToolTip(f"Remove mapping for {field}")
            remove_button.clicked.connect(lambda checked=False, f=field: self.removeRequested.emit(f))
            self.setCellWidget(row, 2, remove_button)

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
        super().keyPressEvent(event)


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
        # Added theme switcher and persistence
        self._apply_theme(self._theme_mode, save=False, update_actions=False)

        self._data_loader = DataLoader()
        self._pdf_engine = PdfEngine()
        self._mapping_manager = MappingManager()

        self._state = UiState()

        self.spreadsheet_panel = SpreadsheetPanel()
        self.pdf_viewer = PdfViewerWidget(self._pdf_engine)
        self.pdf_viewer.fieldAssigned.connect(self._on_field_assigned)
        self.pdf_viewer.pageChanged.connect(self._on_page_changed)
        self.pdf_viewer.zoomChanged.connect(self._on_zoom_changed)

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

        splitter = QtWidgets.QSplitter()
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self.spreadsheet_panel)
        splitter.addWidget(self.viewer_stack)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 4)
        splitter.setSizes([360, 1080])
        self._splitter = splitter

        self.mapping_table = MappingTable()
        self.mapping_table.removeRequested.connect(lambda field: self._action_remove_mapping(field))
        self.mapping_table.itemSelectionChanged.connect(self._update_mapping_action_state)
        self._mapping_dock = QtWidgets.QDockWidget("Mappings", self)
        self._mapping_dock.setWidget(self.mapping_table)
        self._mapping_dock.setAllowedAreas(QtCore.Qt.BottomDockWidgetArea | QtCore.Qt.TopDockWidgetArea)

        container = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(splitter)
        self.setCentralWidget(container)
        self.addDockWidget(QtCore.Qt.BottomDockWidgetArea, self._mapping_dock)

        self._register_actions()
        self._create_menus()
        self._generation_thread: QtCore.QThread | None = None
        self._generation_worker: PdfGenerationWorker | None = None
        self._generation_progress: QtWidgets.QProgressDialog | None = None
        self._configure_page_controls_for_template()
        self._zoom_label = QtWidgets.QLabel("100%")
        self.statusBar().addPermanentWidget(self._zoom_label)
        self._on_zoom_changed(self.pdf_viewer.current_zoom())
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
        self._import_data_action.setIcon(get_fluent_icon("FOLDER"))

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
        self._load_mapping_action.setIcon(get_fluent_icon("FOLDER"))

        self._adjust_range_action = QtGui.QAction("Adjust Data Range", self)
        self._adjust_range_action.setShortcut(QtGui.QKeySequence("Ctrl+Shift+R"))
        # Improved icon contrast with Fluent icons
        self._adjust_range_action.setIcon(get_fluent_icon("SYNC"))
        self._adjust_range_action.triggered.connect(self._action_adjust_data_range)
        self._adjust_range_action.setEnabled(False)

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
                self._remove_mapping_action,
                self._generate_action,
            ]
        )

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
        if not self._state.mapping.assignments:
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
                sheet_msg = f" (sheet '{sample.sheet_name}')" if sample.sheet_name else ""
                self._set_status(
                    f"Loaded data '{sample.source_path.name}'{sheet_msg} ({sample.dataframe.shape[0]:,} rows)"
                )
        else:
            self.spreadsheet_panel.clear()
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

    def _action_generate_pdfs(self) -> None:
        if not self._state.data_sample or not self._state.pdf_template:
            QtWidgets.QMessageBox.warning(
                self, "Missing Data", "Load both a data file and a PDF template first."
            )
            self._set_status("Cannot generate PDFs without data and template", timeout=5000)
            return
        if not self._state.mapping.assignments:
            QtWidgets.QMessageBox.warning(self, "Missing Mapping", "Map at least one field first.")
            self._set_status("Create at least one mapping before generating PDFs", timeout=5000)
            return

        destination = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select Output Directory", ""
        )
        if not destination:
            self._set_status("PDF generation cancelled", timeout=4000)
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

        self._set_status(f"Generating PDFs into '{Path(destination).name}'")
        progress = QtWidgets.QProgressDialog(
            "Generating PDFs...", "Cancel", 0, len(rows), self, QtCore.Qt.WindowTitleHint
        )
        progress.setWindowModality(QtCore.Qt.WindowModal)
        progress.setValue(0)

        worker = PdfGenerationWorker(
            self._pdf_engine,
            self._state.pdf_template.path,
            Path(destination),
            dict(self._state.mapping.assignments),
            rows,
            flatten=True,
            template_metadata=None,
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

        thread.start()
        progress.show()

    # ----- Helpers -------------------------------------------------------------
    def _on_field_assigned(self, field_name: str, column_name: str) -> None:
        self._state.mapping.assign(field_name, column_name)
        self._refresh_mapping_labels()

    def _refresh_mapping_labels(self) -> None:
        self.mapping_table.update_mapping(self._state.mapping.assignments)
        for field_name, column in self._state.mapping.assignments.items():
            preview = self._preview_value_for_column(column)
            self.pdf_viewer.set_assignment(field_name, column, preview)
        self._update_mapping_action_state()

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

    def _preview_value_for_column(self, column_name: str) -> Optional[str]:
        sample = self._state.data_sample
        if not sample or sample.dataframe.empty:
            return None
        frame = sample.dataframe
        if column_name not in frame.columns:
            return None
        matches = np.flatnonzero(frame.columns.to_numpy() == column_name)
        if len(matches) == 0:
            return None
        series = frame.iloc[:, matches[0]]
        if isinstance(series, pd.DataFrame):
            series = series.iloc[:, 0]
        if series.empty:
            return None
        value = series.iloc[0]
        if pd.isna(value):
            return ""
        return str(value)


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
        if hasattr(self, "_remove_mapping_action"):
            self._remove_mapping_action.setEnabled(
                self.mapping_table.selected_field() is not None
            )
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

    def _action_remove_mapping(self, field_name: Optional[str] = None) -> None:
        field = field_name or self.mapping_table.selected_field()
        if not field:
            return
        if field in self._state.mapping.assignments:
            self._state.mapping.remove(field)
            self.pdf_viewer.set_assignment(field, None)
            self._refresh_mapping_labels()
            self._update_data_actions()
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
        destination = outputs[0].parent if outputs else None
        self._cleanup_generation_worker()
        if destination:
            QtWidgets.QMessageBox.information(
                self,
                "Generation Complete",
                f"Created {len(outputs)} PDF files in {destination}.",
            )
            self._set_status(f"Created {len(outputs)} PDF files", timeout=6000)

    def _on_generation_failed(self, message: str) -> None:
        self._cleanup_generation_worker()
        QtWidgets.QMessageBox.critical(self, "Generation Failed", message)
        self._set_status("PDF generation failed", timeout=6000)

    def _on_generation_cancelled(self) -> None:
        self._cleanup_generation_worker()
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
        setThemeColor(ThemeColor.PRIMARY.color())
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





















